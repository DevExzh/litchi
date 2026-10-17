//! The DOC stylesheet and list tables under the DOCX vocabulary (record
//! 0324): `/styles`, `/styles/<id or name>`, `/numbering`,
//! `/numbering/abstractNum[N]`, `/numbering/num[N]` and
//! `/numbering/abstractNum[N]/level[L]`, plus the paragraph and run keys
//! that name a style or a list.
//!
//! Every value is read from the native `STSH`, `PlfLst` and `PlfLfo`
//! structures the open already parsed. A style's properties are its own
//! `UPX` differences, never the inherited chain, as DOCX reads a style's own
//! `w:pPr` and `w:rPr`.

use std::borrow::Cow;
use std::collections::HashSet;

use litchi_doc::parts::chp::{
    CharacterProperties, HighlightColor, UnderlineStyle, VerticalPosition,
};
use litchi_doc::parts::numbering::{
    ListAlignment, ListFollowCharacter, ListLevel, ListStructure, NumberFormat,
};
use litchi_doc::parts::pap::{Justification, ParagraphProperties};
use litchi_doc::parts::styles::{StyleDefinition, StyleKind, StyleSheet};

use super::Doc;
use super::spec;
use crate::node::Value;
use crate::value::{Color, Length};

/// The invariant style identifier Word stores for a user-defined style.
const USER_STYLE: u16 = 0x0FFE;

/// The default character style's fixed index (`Default Paragraph Font`).
pub(super) const DEFAULT_CHARACTER_STYLE: u16 = 10;

/// The DOCX-style identifier of every style slot, by `istd`; empty for an
/// empty slot.
///
/// DOC stores no style identifier, only names, so the identifier follows the
/// rule Word applies when it saves a DOC as DOCX: the name with its spaces
/// and punctuation removed, each word of a built-in style's name
/// capitalized (`heading 1` → `Heading1`). An identifier another style
/// already took gets the style's index appended, so every identifier is
/// unique.
pub(super) fn style_ids(sheet: Option<&StyleSheet>) -> Box<[Box<str>]> {
    let Some(sheet) = sheet else {
        return Box::default();
    };
    let mut seen = HashSet::new();
    sheet
        .styles()
        .iter()
        .enumerate()
        .map(|(index, slot)| {
            let Some(style) = slot else {
                return Box::<str>::default();
            };
            let mut id = base_id(style);
            if id.is_empty() || !seen.insert(id.clone()) {
                id = format!("{id}{index}");
                // An index suffix can collide only with a name that already
                // ends in digits; a second suffix is still unique.
                while !seen.insert(id.clone()) {
                    id.push('_');
                }
            }
            id.into_boxed_str()
        })
        .collect()
}

fn base_id(style: &StyleDefinition) -> String {
    let builtin = style.invariant_id != USER_STYLE;
    let mut id = String::with_capacity(style.name.len());
    for word in style.name.split_whitespace() {
        let mut characters = word.chars().filter(|character| character.is_alphanumeric());
        if let Some(first) = characters.next() {
            if builtin {
                id.extend(first.to_uppercase());
            } else {
                id.push(first);
            }
            id.extend(characters);
        }
    }
    id
}

/// The index of the style a path step names: its identifier, its primary
/// name, or one of its alternate names, each compared exactly.
pub(super) fn find_style(doc: &Doc, name: &str) -> Option<u16> {
    let sheet = doc.parsed().stylesheet()?;
    let ids = doc.style_ids();
    let by_id = ids
        .iter()
        .position(|id| !id.is_empty() && id.as_ref() == name);
    let index = by_id.or_else(|| {
        sheet.styles().iter().position(|slot| {
            slot.as_ref().is_some_and(|style| {
                style.name == name || style.aliases.iter().any(|alias| alias == name)
            })
        })
    })?;
    u16::try_from(index).ok()
}

/// The parsed direct properties of one style, built on its first `format`
/// read.
#[derive(Debug)]
pub(crate) struct StyleView {
    paragraph: Option<ParagraphProperties>,
    character: Option<CharacterProperties>,
}

impl StyleView {
    pub(super) fn of(style: &StyleDefinition) -> Self {
        let paragraph = style
            .paragraph_properties()
            .and_then(|upx| upx.get(2..))
            .and_then(|grpprl| ParagraphProperties::from_sprm(grpprl).ok());
        let character = style
            .character_properties()
            .and_then(|grpprl| CharacterProperties::from_sprm(grpprl).ok());
        Self {
            paragraph,
            character,
        }
    }
}

/// The parsed level paragraph and number properties.
#[derive(Debug)]
pub(crate) struct LevelView {
    paragraph: Option<ParagraphProperties>,
    character: Option<CharacterProperties>,
}

impl LevelView {
    pub(super) fn of(doc: &Doc, definition: usize, level: u8) -> Self {
        let metadata = doc
            .parsed()
            .list_tables()
            .and_then(|tables| tables.metadata().level(definition, level));
        Self {
            paragraph: metadata
                .and_then(|level| ParagraphProperties::from_sprm(&level.paragraph_properties).ok()),
            character: metadata
                .and_then(|level| CharacterProperties::from_sprm(&level.number_properties).ok()),
        }
    }
}

const fn index(table: &[crate::node::spec::PropSpec], name: &str) -> usize {
    let mut at = 0;
    while at < table.len() {
        if same(table[at].name, name) {
            return at;
        }
        at += 1;
    }
    panic!("unknown DOC property");
}

const fn same(left: &str, right: &str) -> bool {
    let (left, right) = (left.as_bytes(), right.as_bytes());
    if left.len() != right.len() {
        return false;
    }
    let mut at = 0;
    while at < left.len() {
        if left[at] != right[at] {
            return false;
        }
        at += 1;
    }
    true
}

macro_rules! keys {
    ($table:expr; $($constant:ident = $name:literal),* $(,)?) => {
        $(const $constant: usize = index($table, $name);)*
    };
}

keys! { spec::STYLE_PROPS;
    S_ID = "id", S_NAME = "name", S_TYPE = "type", S_DEFAULT = "default",
    S_BASED_ON = "basedOn", S_NEXT = "next", S_LINKED = "linked",
    S_AUTO_REDEFINE = "autoRedefine", S_HIDDEN = "hidden", S_QUICK = "qFormat",
    S_PRIORITY = "uiPriority", S_SEMI_HIDDEN = "semiHidden",
    S_UNHIDE = "unhideWhenUsed", S_LOCKED = "locked", S_CUSTOM = "customStyle",
    S_FONT = "font", S_FONT_EA = "font.ea", S_FONT_CS = "font.cs", S_SIZE = "size",
    S_SIZE_CS = "size.cs", S_BOLD = "bold", S_BOLD_CS = "bold.cs", S_ITALIC = "italic",
    S_ITALIC_CS = "italic.cs", S_COLOR = "color", S_UNDERLINE = "underline",
    S_STRIKE = "strike", S_DSTRIKE = "dstrike", S_CAPS = "caps",
    S_SMALL_CAPS = "smallCaps", S_VANISH = "vanish", S_HIGHLIGHT = "highlight",
    S_SUPERSCRIPT = "superscript", S_SUBSCRIPT = "subscript",
    S_CHAR_SPACING = "charSpacing", S_ALIGN = "align", S_SPACE_BEFORE = "spaceBefore",
    S_SPACE_AFTER = "spaceAfter", S_FIRST_LINE = "firstLineIndent",
    S_LEFT = "leftIndent", S_RIGHT = "rightIndent", S_OUTLINE = "outlineLvl",
    S_KEEP_NEXT = "keepNext", S_KEEP_LINES = "keepLines",
    S_PAGE_BREAK = "pageBreakBefore", S_WIDOW = "widowControl",
    S_CONTEXTUAL = "contextualSpacing",
}

keys! { spec::LEVEL_PROPS;
    L_ILVL = "ilvl", L_START = "start", L_FORMAT = "format", L_TEXT = "lvlText",
    L_JUSTIFICATION = "justification", L_SUFFIX = "suff", L_RESTART = "lvlRestart",
    L_LEGAL = "isLgl", L_INDENT = "indent", L_HANGING = "hanging", L_FONT = "font",
    L_SIZE = "size", L_COLOR = "color", L_BOLD = "bold", L_ITALIC = "italic",
}

/// The `ST_StyleType` spelling of a style kind.
const fn style_type(kind: StyleKind) -> &'static str {
    match kind {
        StyleKind::Paragraph => "paragraph",
        StyleKind::Character => "character",
        StyleKind::Table => "table",
        StyleKind::Numbering => "numbering",
    }
}

/// Whether `index` is the fixed default style of its kind (MS-DOC 2.9.274:
/// 0 Normal, 10 Default Paragraph Font, 11 Normal Table, 12 No List).
const fn is_default(style: &StyleDefinition) -> bool {
    matches!(
        (style.index, style.kind),
        (0, StyleKind::Paragraph)
            | (10, StyleKind::Character)
            | (11, StyleKind::Table)
            | (12, StyleKind::Numbering)
    )
}

fn font<'a>(doc: &'a Doc, ftc: Option<u16>) -> Option<Value<'a>> {
    let name = doc.parsed().font_name(ftc?).ok()??;
    Some(Value::Text(Cow::Borrowed(name)))
}

fn half_points<'a>(size: Option<u16>) -> Option<Value<'a>> {
    size.map(|size| Value::Length(Length::from_half_points(i32::from(size))))
}

fn twips<'a>(value: Option<i32>) -> Option<Value<'a>> {
    value.map(|value| Value::Length(Length::from_twips(value)))
}

/// The value of character key `index` of a style or level table.
fn character_value<'a>(
    doc: &'a Doc,
    character: &CharacterProperties,
    key: CharacterKey,
) -> Option<Value<'a>> {
    match key {
        CharacterKey::Font => font(doc, character.font_index),
        CharacterKey::FontEastAsia => font(doc, character.font_index_fe),
        CharacterKey::FontComplexScript => font(doc, character.font_index_bidi),
        CharacterKey::Size => half_points(character.font_size),
        CharacterKey::SizeComplexScript => half_points(character.font_size_bidi),
        CharacterKey::Bold => character.is_bold.map(Value::Bool),
        CharacterKey::BoldComplexScript => character.is_bold_bidi.map(Value::Bool),
        CharacterKey::Italic => character.is_italic.map(Value::Bool),
        CharacterKey::ItalicComplexScript => character.is_italic_bidi.map(Value::Bool),
        CharacterKey::Color => character
            .color
            .map(|(red, green, blue)| Value::Color(Color::rgb(red, green, blue))),
        CharacterKey::Underline => (character.underline != UnderlineStyle::None)
            .then(|| Value::Token(spec::underline_name(character.underline))),
        CharacterKey::Strike => character.is_strikethrough.map(Value::Bool),
        CharacterKey::DoubleStrike => character.is_double_strikethrough.map(Value::Bool),
        CharacterKey::Caps => character.is_all_caps.map(Value::Bool),
        CharacterKey::SmallCaps => character.is_small_caps.map(Value::Bool),
        CharacterKey::Vanish => character.is_hidden.map(Value::Bool),
        CharacterKey::Highlight => character
            .highlight
            .map(|color| Value::Token(highlight(color))),
        CharacterKey::Superscript => (character.vertical_position == VerticalPosition::Superscript)
            .then_some(Value::Bool(true)),
        CharacterKey::Subscript => (character.vertical_position == VerticalPosition::Subscript)
            .then_some(Value::Bool(true)),
        CharacterKey::Spacing => twips(character.char_spacing.map(i32::from)),
    }
}

#[derive(Clone, Copy)]
enum CharacterKey {
    Font,
    FontEastAsia,
    FontComplexScript,
    Size,
    SizeComplexScript,
    Bold,
    BoldComplexScript,
    Italic,
    ItalicComplexScript,
    Color,
    Underline,
    Strike,
    DoubleStrike,
    Caps,
    SmallCaps,
    Vanish,
    Highlight,
    Superscript,
    Subscript,
    Spacing,
}

/// The `ST_HighlightColor` spelling of a highlight.
pub(super) const fn highlight(color: HighlightColor) -> &'static str {
    use HighlightColor as H;
    match color {
        H::None => "none",
        H::Black => "black",
        H::Blue => "blue",
        H::Cyan => "cyan",
        H::Green => "green",
        H::Magenta => "magenta",
        H::Red => "red",
        H::Yellow => "yellow",
        H::White => "white",
        H::DarkBlue => "darkBlue",
        H::DarkCyan => "darkCyan",
        H::DarkGreen => "darkGreen",
        H::DarkMagenta => "darkMagenta",
        H::DarkRed => "darkRed",
        H::DarkYellow => "darkYellow",
        H::DarkGray => "darkGray",
        H::LightGray => "lightGray",
    }
}

/// A style's `id`, the cached identifier of `istd`.
fn style_id(doc: &Doc, istd: u16) -> Option<Value<'_>> {
    doc.style_id(istd)
        .filter(|id| !id.is_empty())
        .map(|id| Value::Text(Cow::Borrowed(id)))
}

/// The value at `index` of the style table for style `istd`.
pub(super) fn style_value<'a>(
    doc: &'a Doc,
    istd: u16,
    view: &StyleView,
    index: usize,
) -> Option<Value<'a>> {
    let sheet = doc.parsed().stylesheet()?;
    let style = sheet.get(istd)?;
    let character = |key| {
        view.character
            .as_ref()
            .and_then(|character| character_value(doc, character, key))
    };
    let paragraph = view.paragraph.as_ref();
    let post = style.post_2000.as_ref();
    match index {
        S_ID => style_id(doc, istd),
        S_NAME => Some(Value::Text(Cow::Borrowed(style.name.as_str()))),
        S_TYPE => Some(Value::Token(style_type(style.kind))),
        S_DEFAULT => is_default(style).then_some(Value::Bool(true)),
        S_BASED_ON => style.base_style.and_then(|base| style_id(doc, base)),
        S_NEXT => (style.kind == StyleKind::Paragraph && sheet.get(style.next_style).is_some())
            .then(|| style_id(doc, style.next_style))
            .flatten(),
        S_LINKED => post
            .and_then(|post| post.linked_style)
            .and_then(|linked| style_id(doc, linked)),
        S_AUTO_REDEFINE => style.flags.auto_redefine.then_some(Value::Bool(true)),
        S_HIDDEN => style.flags.hidden.then_some(Value::Bool(true)),
        S_QUICK => style.flags.quick_format.then_some(Value::Bool(true)),
        S_PRIORITY => post.map(|post| Value::Int(i64::from(post.priority))),
        S_SEMI_HIDDEN => style.flags.semi_hidden.then_some(Value::Bool(true)),
        S_UNHIDE => style.flags.unhide_when_used.then_some(Value::Bool(true)),
        S_LOCKED => style.flags.locked.then_some(Value::Bool(true)),
        S_CUSTOM => (style.invariant_id == USER_STYLE).then_some(Value::Bool(true)),
        S_FONT => character(CharacterKey::Font),
        S_FONT_EA => character(CharacterKey::FontEastAsia),
        S_FONT_CS => character(CharacterKey::FontComplexScript),
        S_SIZE => character(CharacterKey::Size),
        S_SIZE_CS => character(CharacterKey::SizeComplexScript),
        S_BOLD => character(CharacterKey::Bold),
        S_BOLD_CS => character(CharacterKey::BoldComplexScript),
        S_ITALIC => character(CharacterKey::Italic),
        S_ITALIC_CS => character(CharacterKey::ItalicComplexScript),
        S_COLOR => character(CharacterKey::Color),
        S_UNDERLINE => character(CharacterKey::Underline),
        S_STRIKE => character(CharacterKey::Strike),
        S_DSTRIKE => character(CharacterKey::DoubleStrike),
        S_CAPS => character(CharacterKey::Caps),
        S_SMALL_CAPS => character(CharacterKey::SmallCaps),
        S_VANISH => character(CharacterKey::Vanish),
        S_HIGHLIGHT => character(CharacterKey::Highlight),
        S_SUPERSCRIPT => character(CharacterKey::Superscript),
        S_SUBSCRIPT => character(CharacterKey::Subscript),
        S_CHAR_SPACING => character(CharacterKey::Spacing),
        S_ALIGN => paragraph
            .filter(|paragraph| paragraph.justification != Justification::Left)
            .map(|paragraph| Value::Token(spec::alignment_name(paragraph.justification))),
        S_SPACE_BEFORE => paragraph
            .and_then(|paragraph| paragraph.space_before)
            .and_then(|value| twips(Some(i32::from(value)))),
        S_SPACE_AFTER => paragraph
            .and_then(|paragraph| paragraph.space_after)
            .and_then(|value| twips(Some(i32::from(value)))),
        S_FIRST_LINE => twips(paragraph.and_then(|paragraph| paragraph.indent_first_line)),
        S_LEFT => twips(paragraph.and_then(|paragraph| paragraph.indent_left)),
        S_RIGHT => twips(paragraph.and_then(|paragraph| paragraph.indent_right)),
        S_OUTLINE => paragraph
            .and_then(|paragraph| paragraph.outline_level)
            .filter(|level| *level < 9)
            .map(|level| Value::Int(i64::from(level))),
        S_KEEP_NEXT => paragraph
            .filter(|paragraph| paragraph.keep_with_next)
            .map(|_| Value::Bool(true)),
        S_KEEP_LINES => paragraph
            .filter(|paragraph| paragraph.keep_on_page)
            .map(|_| Value::Bool(true)),
        S_PAGE_BREAK => paragraph
            .filter(|paragraph| paragraph.page_break_before)
            .map(|_| Value::Bool(true)),
        S_WIDOW => paragraph
            .filter(|paragraph| paragraph.widow_control)
            .map(|_| Value::Bool(true)),
        S_CONTEXTUAL => paragraph
            .filter(|paragraph| paragraph.contextual_spacing)
            .map(|_| Value::Bool(true)),
        _ => None,
    }
}

/// The `ST_NumberFormat` spelling of a DOC number format (`MSONFC`).
pub(crate) const fn number_format(format: NumberFormat) -> &'static str {
    use NumberFormat as F;
    match format {
        F::Arabic => "decimal",
        F::UpperRoman => "upperRoman",
        F::LowerRoman => "lowerRoman",
        F::UpperLetter => "upperLetter",
        F::LowerLetter => "lowerLetter",
        F::Ordinal => "ordinal",
        F::CardinalText => "cardinalText",
        F::OrdinalText => "ordinalText",
        F::Hex => "hex",
        F::Chicago => "chicago",
        F::IdeographDigital => "ideographDigital",
        F::JapaneseCounting => "japaneseCounting",
        F::Aiueo => "aiueo",
        F::Iroha => "iroha",
        F::DecimalFullWidth => "decimalFullWidth",
        F::DecimalHalfWidth => "decimalHalfWidth",
        F::JapaneseLegal => "japaneseLegal",
        F::JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
        F::DecimalEnclosedCircle => "decimalEnclosedCircle",
        F::DecimalFullWidth2 => "decimalFullWidth2",
        F::AiueoFullWidth => "aiueoFullWidth",
        F::IrohaFullWidth => "irohaFullWidth",
        F::DecimalZero => "decimalZero",
        F::Bullet => "bullet",
        F::Ganada => "ganada",
        F::Chosung => "chosung",
        F::DecimalEnclosedFullstop => "decimalEnclosedFullstop",
        F::DecimalEnclosedParen => "decimalEnclosedParen",
        F::DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
        F::IdeographEnclosedCircle => "ideographEnclosedCircle",
        F::IdeographTraditional => "ideographTraditional",
        F::IdeographZodiac => "ideographZodiac",
        F::IdeographZodiacTraditional => "ideographZodiacTraditional",
        F::TaiwaneseCounting => "taiwaneseCounting",
        F::IdeographLegalTraditional => "ideographLegalTraditional",
        F::TaiwaneseCountingThousand => "taiwaneseCountingThousand",
        F::TaiwaneseDigital => "taiwaneseDigital",
        F::ChineseCounting => "chineseCounting",
        F::ChineseLegalSimplified => "chineseLegalSimplified",
        F::ChineseCountingThousand => "chineseCountingThousand",
        F::DecimalChinese => "decimal",
        F::KoreanDigital => "koreanDigital",
        F::KoreanCounting => "koreanCounting",
        F::KoreanLegal => "koreanLegal",
        F::KoreanDigital2 => "koreanDigital2",
        F::Hebrew1 => "hebrew1",
        F::ArabicAlpha => "arabicAlpha",
        F::Hebrew2 => "hebrew2",
        F::ArabicAbjad => "arabicAbjad",
        F::HindiVowels => "hindiVowels",
        F::HindiConsonants => "hindiConsonants",
        F::HindiNumbers => "hindiNumbers",
        F::HindiCounting => "hindiCounting",
        F::ThaiLetters => "thaiLetters",
        F::ThaiNumbers => "thaiNumbers",
        F::ThaiCounting => "thaiCounting",
        F::VietnameseCounting => "vietnameseCounting",
        F::NumberInDash => "numberInDash",
        F::RussianLower => "russianLower",
        F::RussianUpper => "russianUpper",
        F::None => "none",
    }
}

/// The list definitions of the document.
pub(super) fn structures(doc: &Doc) -> &[ListStructure] {
    doc.reader()
        .ok()
        .and_then(|reader| reader.list_tables())
        .map_or(&[], |tables| tables.structures())
}

/// The number of list instances of the document.
pub(super) fn num_count(doc: &Doc) -> usize {
    doc.reader()
        .ok()
        .and_then(|reader| reader.list_tables())
        .map_or(0, |tables| tables.overrides().len())
}

/// The value at `index` of the definition table for definition `at`.
pub(super) fn abstract_num_value(doc: &Doc, at: usize, index: usize) -> Option<Value<'_>> {
    let reader = doc.parsed();
    let tables = reader.list_tables()?;
    let structure = tables.structures().get(at)?;
    match index {
        0 => i64::try_from(at).ok().map(Value::Int),
        1 => {
            let hybrid = tables
                .metadata()
                .definition(at)
                .is_some_and(|definition| definition.hybrid);
            Some(Value::Token(if hybrid {
                "hybridMultilevel"
            } else if structure.is_simple {
                "singleLevel"
            } else {
                "multilevel"
            }))
        },
        2 => reader
            .list_name_for_definition_index(at)
            .filter(|name| !name.is_empty())
            .map(|name| Value::Text(Cow::Borrowed(name))),
        _ => None,
    }
}

/// The value at `index` of the instance table for instance `at`.
pub(super) fn num_value(doc: &Doc, at: usize, index: usize) -> Option<Value<'_>> {
    let tables = doc.parsed().list_tables()?;
    let instance = tables.overrides().get(at)?;
    match index {
        0 => i64::try_from(at + 1).ok().map(Value::Int),
        1 => tables
            .structures()
            .iter()
            .position(|structure| structure.list_id == instance.list_id)
            .and_then(|position| i64::try_from(position).ok())
            .map(Value::Int),
        level @ 2..=10 => {
            let level = u8::try_from(level - 2).ok()?;
            instance
                .level_override(level)
                .and_then(|lfolvl| lfolvl.start_at)
                .map(|start| Value::Int(i64::from(start)))
        },
        _ => None,
    }
}

/// The value at `index` of the level table for `level` of definition
/// `definition`.
pub(super) fn level_value<'a>(
    doc: &'a Doc,
    definition: usize,
    level: &'a ListLevel,
    view: &LevelView,
    index: usize,
) -> Option<Value<'a>> {
    let metadata = doc
        .parsed()
        .list_tables()
        .and_then(|tables| tables.metadata().level(definition, level.level));
    let character = |key| {
        view.character
            .as_ref()
            .and_then(|character| character_value(doc, character, key))
    };
    let paragraph = view.paragraph.as_ref();
    match index {
        L_ILVL => Some(Value::Int(i64::from(level.level))),
        L_START => Some(Value::Int(i64::from(level.start_at))),
        L_FORMAT => Some(Value::Text(Cow::Borrowed(number_format(
            level.number_format,
        )))),
        L_TEXT => Some(Value::Text(Cow::Borrowed(level.number_text.as_str()))),
        L_JUSTIFICATION => Some(Value::Token(match level.alignment {
            ListAlignment::Left => "left",
            ListAlignment::Center => "center",
            ListAlignment::Right => "right",
        })),
        L_SUFFIX => metadata.map(|metadata| {
            Value::Token(match metadata.follow_character {
                ListFollowCharacter::Tab => "tab",
                ListFollowCharacter::Space => "space",
                ListFollowCharacter::Nothing => "nothing",
            })
        }),
        L_RESTART => metadata
            .and_then(|metadata| metadata.restart_limit)
            .map(|limit| Value::Int(i64::from(limit))),
        L_LEGAL => metadata
            .filter(|metadata| metadata.legal_numbering)
            .map(|_| Value::Bool(true)),
        L_INDENT => paragraph
            .and_then(|paragraph| paragraph.indent_left)
            .map(|value| Value::Int(i64::from(value))),
        L_HANGING => paragraph
            .and_then(|paragraph| paragraph.indent_first_line)
            .filter(|value| *value < 0)
            .map(|value| Value::Int(-i64::from(value))),
        L_FONT => character(CharacterKey::Font),
        L_SIZE => character(CharacterKey::Size),
        L_COLOR => character(CharacterKey::Color),
        L_BOLD => character(CharacterKey::Bold),
        L_ITALIC => character(CharacterKey::Italic),
        _ => None,
    }
}

/// A paragraph's `style`, `styleId` or `styleName`, or its list `numFmt`,
/// `listStyle`, `start` or `numLevel`. The `Normal` style (`istd` 0) is
/// reported as no style, as a DOCX paragraph without `w:pStyle` is.
pub(super) fn paragraph_value<'a>(
    doc: &'a Doc,
    paragraph: &litchi_doc::Paragraph,
    key: ParagraphKey,
) -> Option<Value<'a>> {
    let reader = doc.parsed();
    match key {
        ParagraphKey::StyleId => {
            let istd = paragraph
                .properties()
                .style_index
                .filter(|istd| *istd != 0)?;
            style_id(doc, istd)
        },
        ParagraphKey::StyleName => {
            let istd = paragraph
                .properties()
                .style_index
                .filter(|istd| *istd != 0)?;
            let style = reader.stylesheet()?.get(istd)?;
            Some(Value::Text(Cow::Borrowed(style.name.as_str())))
        },
        ParagraphKey::NumberFormat
        | ParagraphKey::ListStyle
        | ParagraphKey::Start
        | ParagraphKey::Level => {
            let binding = reader.paragraph_list_binding(paragraph)?;
            let level = binding.effective_level();
            match key {
                ParagraphKey::NumberFormat => Some(Value::Text(Cow::Borrowed(number_format(
                    level.number_format,
                )))),
                ParagraphKey::ListStyle => Some(Value::Token(if level.is_bullet() {
                    "bullet"
                } else {
                    "ordered"
                })),
                ParagraphKey::Start => Some(Value::Int(i64::from(binding.effective_start_at()))),
                _ => Some(Value::Int(i64::from(binding.level))),
            }
        },
    }
}

/// The paragraph keys [`paragraph_value`] answers.
#[derive(Clone, Copy)]
pub(super) enum ParagraphKey {
    StyleId,
    StyleName,
    NumberFormat,
    ListStyle,
    Start,
    Level,
}

/// A run's `rStyle`: its character style, unless that is the default
/// character style.
pub(super) fn run_style<'a>(doc: &'a Doc, run: &litchi_doc::Run) -> Option<Value<'a>> {
    let istd = run
        .properties()
        .style_index
        .filter(|istd| *istd != DEFAULT_CHARACTER_STYLE)?;
    style_id(doc, istd)
}
