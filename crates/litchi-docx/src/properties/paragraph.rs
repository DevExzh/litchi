//! Direct paragraph properties: the children of a paragraph's own `w:pPr`.

use quick_xml::events::BytesStart;

use super::run::{self, toggle, value};
use super::scan::{self, Context, Site};
use super::value::{Border, Shading, Sign, Twips, Unit, integer, measure, on_off, vocabulary};
use super::{Invalid, collect, optional, parse_border, parse_shading};
use crate::error::Result;
use crate::paragraph::{LineSpacingRule, ParagraphSpacing};

/// The paragraph-property elements this reader types, in schema order. A
/// property's position here is its bit in
/// [`Paragraph::invalid_properties`].
const ELEMENTS: [&str; 31] = [
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
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
    "rPr",
];

const _: () = assert!(ELEMENTS.len() <= 64, "an Invalid mask holds 64 properties");

const NUMBERING: usize = super::position(&ELEMENTS, "numPr");
const BORDERS: usize = super::position(&ELEMENTS, "pBdr");
const TABS: usize = super::position(&ELEMENTS, "tabs");

vocabulary! {
    /// Paragraph alignment (`ST_Jc`). `left` and `right` are the
    /// transitional spellings of `start` and `end`; `both` justifies.
    pub enum Justification {
        /// `start`
        Start = "start",
        /// `center`
        Center = "center",
        /// `end`
        End = "end",
        /// `both`: justified.
        Both = "both",
        /// `mediumKashida`
        MediumKashida = "mediumKashida",
        /// `distribute`
        Distribute = "distribute",
        /// `numTab`
        NumTab = "numTab",
        /// `highKashida`
        HighKashida = "highKashida",
        /// `lowKashida`
        LowKashida = "lowKashida",
        /// `thaiDistribute`
        ThaiDistribute = "thaiDistribute",
        /// `left`
        Left = "left",
        /// `right`
        Right = "right",
    }
}

vocabulary! {
    /// The vertical alignment of characters on a line (`ST_TextAlignment`).
    pub enum TextAlignment {
        /// `top`
        Top = "top",
        /// `center`
        Center = "center",
        /// `baseline`
        Baseline = "baseline",
        /// `bottom`
        Bottom = "bottom",
        /// `auto`
        Auto = "auto",
    }
}

vocabulary! {
    /// Text flow direction (`ST_TextDirection`) of a paragraph or cell. The
    /// last six values are the transitional spellings.
    pub enum TextDirection {
        /// `tb`: top to bottom, left to right.
        Tb = "tb",
        /// `rl`: right to left, top to bottom.
        Rl = "rl",
        /// `lr`: left to right, bottom to top.
        Lr = "lr",
        /// `tbV`
        TbV = "tbV",
        /// `rlV`
        RlV = "rlV",
        /// `lrV`
        LrV = "lrV",
        /// `btLr`
        BtLr = "btLr",
        /// `lrTb`
        LrTb = "lrTb",
        /// `lrTbV`
        LrTbV = "lrTbV",
        /// `tbLrV`
        TbLrV = "tbLrV",
        /// `tbRl`
        TbRl = "tbRl",
        /// `tbRlV`
        TbRlV = "tbRlV",
    }
}

vocabulary! {
    /// How a text box's lines wrap tightly (`ST_TextboxTightWrap`).
    pub enum TextboxTightWrap {
        /// `none`
        None = "none",
        /// `allLines`
        AllLines = "allLines",
        /// `firstAndLastLine`
        FirstAndLastLine = "firstAndLastLine",
        /// `firstLineOnly`
        FirstLineOnly = "firstLineOnly",
        /// `lastLineOnly`
        LastLineOnly = "lastLineOnly",
    }
}

vocabulary! {
    /// A tab stop's alignment (`ST_TabJc`). `clear` removes an inherited
    /// stop; `left` and `right` are the transitional spellings of `start`
    /// and `end`.
    pub enum TabKind {
        /// `clear`
        Clear = "clear",
        /// `start`
        Start = "start",
        /// `center`
        Center = "center",
        /// `end`
        End = "end",
        /// `decimal`
        Decimal = "decimal",
        /// `bar`
        Bar = "bar",
        /// `num`
        Num = "num",
        /// `left`
        Left = "left",
        /// `right`
        Right = "right",
    }
}

vocabulary! {
    /// A tab stop's leader (`ST_TabTlc`).
    pub enum TabLeader {
        /// `none`
        None = "none",
        /// `dot`
        Dot = "dot",
        /// `hyphen`
        Hyphen = "hyphen",
        /// `underscore`
        Underscore = "underscore",
        /// `heavy`
        Heavy = "heavy",
        /// `middleDot`
        MiddleDot = "middleDot",
    }
}

/// One direct tab stop (`w:tabs/w:tab`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TabStop {
    /// `w:val`
    pub kind: TabKind,
    /// `w:leader`
    pub leader: Option<TabLeader>,
    /// `w:pos`
    pub position: Twips,
}

impl TabStop {
    /// A tab stop of `kind` at `position`, without a leader.
    ///
    /// ```
    /// use litchi_docx::properties::{TabKind, TabLeader, TabStop, Twips};
    ///
    /// let tab = TabStop::new(TabKind::Decimal, Twips(1440)).with_leader(TabLeader::Dot);
    /// assert_eq!(tab.leader, Some(TabLeader::Dot));
    /// ```
    #[must_use]
    pub const fn new(kind: TabKind, position: Twips) -> Self {
        Self {
            kind,
            leader: None,
            position,
        }
    }

    /// The same tab stop with `leader`.
    #[must_use]
    pub const fn with_leader(mut self, leader: TabLeader) -> Self {
        self.leader = Some(leader);
        self
    }
}

/// Direct paragraph numbering (`w:numPr`), as authored.
///
/// An empty `w:numPr`, or one with only a level, names no numbering
/// instance: the level then overrides the one inherited from the style.
/// [`crate::Paragraph::numbering`] resolves these cases.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Numbering {
    /// `w:numId`; `0` removes numbering inherited from the style.
    pub id: Option<u32>,
    /// `w:ilvl`, as authored (Word uses levels 0 to 8).
    pub level: Option<u8>,
}

/// Direct paragraph indentation (`w:ind`).
///
/// `start`/`end` and the transitional `left`/`right` are kept apart as
/// authored; Word reads `left` as `start` and `right` as `end`, and
/// [`Self::leading`] and [`Self::trailing`] apply that. The `*_chars` values
/// are hundredths of a character width and take precedence over the
/// matching length.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Indentation {
    /// `w:start`
    pub start: Option<Twips>,
    /// `w:startChars`
    pub start_chars: Option<i64>,
    /// `w:end`
    pub end: Option<Twips>,
    /// `w:endChars`
    pub end_chars: Option<i64>,
    /// `w:left`
    pub left: Option<Twips>,
    /// `w:leftChars`
    pub left_chars: Option<i64>,
    /// `w:right`
    pub right: Option<Twips>,
    /// `w:rightChars`
    pub right_chars: Option<i64>,
    /// `w:hanging`
    pub hanging: Option<Twips>,
    /// `w:hangingChars`
    pub hanging_chars: Option<i64>,
    /// `w:firstLine`
    pub first_line: Option<Twips>,
    /// `w:firstLineChars`
    pub first_line_chars: Option<i64>,
}

impl Indentation {
    /// The leading indent: `w:left`, else `w:start`.
    #[must_use]
    pub fn leading(&self) -> Option<Twips> {
        self.left.or(self.start)
    }

    /// The trailing indent: `w:right`, else `w:end`.
    #[must_use]
    pub fn trailing(&self) -> Option<Twips> {
        self.right.or(self.end)
    }
}

/// Direct paragraph borders (`w:pBdr`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct ParagraphBorders {
    /// `w:top`
    pub top: Option<Border>,
    /// `w:left`
    pub left: Option<Border>,
    /// `w:bottom`
    pub bottom: Option<Border>,
    /// `w:right`
    pub right: Option<Border>,
    /// `w:between`: between paragraphs with the same borders.
    pub between: Option<Border>,
    /// `w:bar`: a vertical bar beside the paragraph.
    pub bar: Option<Border>,
}

/// Direct (local) paragraph properties: what one paragraph's own `w:pPr`
/// states.
///
/// Read with [`crate::Paragraph::direct_properties`]. Every field is exactly
/// what the paragraph authors; `None` means the paragraph does not state the
/// property, which is then inherited from its style and the document
/// defaults. The paragraph mark's own run properties are in [`Self::mark`].
/// The previous formatting of a tracked change (`w:pPrChange`) and a section
/// break's properties (`w:sectPr`) are not read.
///
/// ```
/// use litchi_docx::Paragraph;
/// use litchi_docx::properties::{Justification, Twips};
///
/// let paragraph = Paragraph::new(
///     br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/><w:jc w:val="center"/><w:ind w:left="720" w:hanging="360"/><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t>Title</w:t></w:r></w:p>"#.to_vec(),
/// );
/// let properties = paragraph.direct_properties()?;
/// assert_eq!(properties.style.as_deref(), Some("Heading1"));
/// assert_eq!(properties.keep_next, Some(true));
/// assert_eq!(properties.justification, Some(Justification::Center));
/// let indentation = properties.indentation.unwrap();
/// assert_eq!(indentation.leading(), Some(Twips(720)));
/// assert_eq!(indentation.hanging, Some(Twips(360)));
/// assert_eq!(properties.outline_level, Some(0));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Paragraph {
    /// `w:pStyle`: the paragraph style identifier.
    pub style: Option<String>,
    /// `w:keepNext`: keep on the same page as the next paragraph.
    pub keep_next: Option<bool>,
    /// `w:keepLines`: keep all lines on one page.
    pub keep_lines: Option<bool>,
    /// `w:pageBreakBefore`
    pub page_break_before: Option<bool>,
    /// `w:widowControl`
    pub widow_control: Option<bool>,
    /// `w:numPr`
    pub numbering: Option<Numbering>,
    /// `w:suppressLineNumbers`
    pub suppress_line_numbers: Option<bool>,
    /// `w:pBdr`
    pub borders: Option<ParagraphBorders>,
    /// `w:shd`
    pub shading: Option<Shading>,
    /// `w:tabs`, in document order; an empty `w:tabs` is `Some` and empty.
    pub tabs: Option<Vec<TabStop>>,
    /// `w:suppressAutoHyphens`
    pub suppress_auto_hyphens: Option<bool>,
    /// `w:kinsoku`: East Asian line-breaking rules.
    pub kinsoku: Option<bool>,
    /// `w:wordWrap`
    pub word_wrap: Option<bool>,
    /// `w:overflowPunct`
    pub overflow_punctuation: Option<bool>,
    /// `w:topLinePunct`
    pub top_line_punctuation: Option<bool>,
    /// `w:autoSpaceDE`: space between East Asian and Latin text.
    pub auto_space_east_asian_latin: Option<bool>,
    /// `w:autoSpaceDN`: space between East Asian text and numbers.
    pub auto_space_east_asian_numbers: Option<bool>,
    /// `w:bidi`: a right-to-left paragraph.
    pub bidi: Option<bool>,
    /// `w:adjustRightInd`
    pub adjust_right_indent: Option<bool>,
    /// `w:snapToGrid`
    pub snap_to_grid: Option<bool>,
    /// `w:spacing`, typed as [`crate::Paragraph::spacing`] types it; a
    /// universal measure or decimal is converted to twips here, where that
    /// reader refuses it.
    pub spacing: Option<ParagraphSpacing>,
    /// `w:ind`
    pub indentation: Option<Indentation>,
    /// `w:contextualSpacing`: no spacing between paragraphs of one style.
    pub contextual_spacing: Option<bool>,
    /// `w:mirrorIndents`
    pub mirror_indents: Option<bool>,
    /// `w:suppressOverlap`
    pub suppress_overlap: Option<bool>,
    /// `w:jc`
    pub justification: Option<Justification>,
    /// `w:textDirection`
    pub text_direction: Option<TextDirection>,
    /// `w:textAlignment`
    pub text_alignment: Option<TextAlignment>,
    /// `w:textboxTightWrap`
    pub textbox_tight_wrap: Option<TextboxTightWrap>,
    /// `w:outlineLvl`: 0 to 8 are heading levels and 9 is body text.
    pub outline_level: Option<u8>,
    /// `w:rPr`: the run properties of the paragraph mark.
    pub mark: Option<Box<run::Run>>,
    pub(super) invalid: Invalid,
}

impl Paragraph {
    /// The direct properties the paragraph states with a value outside the
    /// property's schema domain, by element local name in schema order.
    ///
    /// Each such property is left unset rather than guessed; for `w:ind`
    /// only the invalid attribute is. A tab stop or border that cannot be
    /// typed is left out of [`Self::tabs`] or [`Self::borders`]. The
    /// paragraph mark's invalid properties are listed by
    /// [`Self::mark`]'s own [`run::Run::invalid_properties`].
    pub fn invalid_properties(&self) -> impl Iterator<Item = &'static str> + '_ {
        self.invalid.names(&ELEMENTS)
    }
}

/// Collects one paragraph's direct properties; repeated elements are read
/// from their first occurrence, and `w:ind` attributes from the first
/// occurrence that states each.
#[derive(Default)]
struct Reader {
    paragraph: Paragraph,
    seen: u64,
    mark: Option<run::Reader>,
    /// Whether the property element whose children are being read is the
    /// first of its name.
    first: bool,
}

impl Reader {
    fn invalid(&mut self, index: usize) {
        self.paragraph.invalid.insert(index);
    }

    fn property(
        &mut self,
        context: &Context<'_>,
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        let Some(index) = ELEMENTS.iter().position(|name| name.as_bytes() == local) else {
            self.first = false;
            return Ok(());
        };
        let bit = 1u64 << index;
        let repeated = self.seen & bit != 0;
        self.seen |= bit;
        self.first = !repeated;
        let paragraph = &mut self.paragraph;
        let valid = match local {
            b"ind" => {
                let [
                    start,
                    start_chars,
                    end,
                    end_chars,
                    left,
                    left_chars,
                    right,
                    right_chars,
                    hanging,
                    hanging_chars,
                    first_line,
                    first_line_chars,
                ] = collect(
                    context,
                    element,
                    [
                        b"start",
                        b"startChars",
                        b"end",
                        b"endChars",
                        b"left",
                        b"leftChars",
                        b"right",
                        b"rightChars",
                        b"hanging",
                        b"hangingChars",
                        b"firstLine",
                        b"firstLineChars",
                    ],
                )?;
                let indentation = paragraph
                    .indentation
                    .get_or_insert_with(Indentation::default);
                let mut valid = true;
                let signed = |raw: &str| measure(raw, Unit::Twips, Sign::Signed).map(Twips);
                let unsigned = |raw: &str| measure(raw, Unit::Twips, Sign::Unsigned).map(Twips);
                fill(&mut indentation.start, start.as_deref(), signed, &mut valid);
                fill(&mut indentation.end, end.as_deref(), signed, &mut valid);
                fill(&mut indentation.left, left.as_deref(), signed, &mut valid);
                fill(&mut indentation.right, right.as_deref(), signed, &mut valid);
                fill(
                    &mut indentation.hanging,
                    hanging.as_deref(),
                    unsigned,
                    &mut valid,
                );
                fill(
                    &mut indentation.first_line,
                    first_line.as_deref(),
                    unsigned,
                    &mut valid,
                );
                fill(
                    &mut indentation.start_chars,
                    start_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                fill(
                    &mut indentation.end_chars,
                    end_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                fill(
                    &mut indentation.left_chars,
                    left_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                fill(
                    &mut indentation.right_chars,
                    right_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                fill(
                    &mut indentation.hanging_chars,
                    hanging_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                fill(
                    &mut indentation.first_line_chars,
                    first_line_chars.as_deref(),
                    integer,
                    &mut valid,
                );
                valid
            },
            _ if repeated => true,
            b"pStyle" => {
                let [value] = collect(context, element, [b"val"])?;
                value
                    .map(|value| paragraph.style = Some(value.into_owned()))
                    .is_some()
            },
            b"numPr" => {
                paragraph.numbering = Some(Numbering::default());
                true
            },
            b"pBdr" => {
                paragraph.borders = Some(ParagraphBorders::default());
                true
            },
            b"tabs" => {
                paragraph.tabs = Some(Vec::new());
                true
            },
            b"rPr" => {
                self.mark = Some(run::Reader::default());
                true
            },
            b"shd" => set(&mut paragraph.shading, parse_shading(context, element)?),
            b"spacing" => set(&mut paragraph.spacing, parse_spacing(context, element)?),
            b"jc" => set(
                &mut paragraph.justification,
                value(context, element, Justification::from_xml)?,
            ),
            b"textDirection" => set(
                &mut paragraph.text_direction,
                value(context, element, TextDirection::from_xml)?,
            ),
            b"textAlignment" => set(
                &mut paragraph.text_alignment,
                value(context, element, TextAlignment::from_xml)?,
            ),
            b"textboxTightWrap" => set(
                &mut paragraph.textbox_tight_wrap,
                value(context, element, TextboxTightWrap::from_xml)?,
            ),
            b"outlineLvl" => set(
                &mut paragraph.outline_level,
                value(context, element, |raw| {
                    integer(raw)
                        .and_then(|level| u8::try_from(level).ok())
                        .filter(|level| *level <= 9)
                })?,
            ),
            _ => {
                let toggle = toggle(context, element)?;
                let Some(slot) = toggle_slot(paragraph, local) else {
                    return Ok(());
                };
                set(slot, toggle)
            },
        };
        if !valid {
            self.invalid(index);
        }
        Ok(())
    }

    /// Read a direct child of a `w:numPr`, `w:pBdr`, `w:tabs` or `w:rPr`.
    fn nested(
        &mut self,
        context: &Context<'_>,
        parent: &[u8],
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if !self.first {
            return Ok(());
        }
        let (index, valid) = match parent {
            b"rPr" => {
                if let Some(mark) = self.mark.as_mut() {
                    mark.property(context, local, element)?;
                }
                return Ok(());
            },
            b"numPr" => {
                let Some(numbering) = self.paragraph.numbering.as_mut() else {
                    return Ok(());
                };
                let valid = match local {
                    b"ilvl" => fill_value(
                        &mut numbering.level,
                        value(context, element, |raw| {
                            integer(raw).and_then(|level| u8::try_from(level).ok())
                        })?,
                    ),
                    b"numId" => fill_value(
                        &mut numbering.id,
                        value(context, element, super::value::unsigned)?,
                    ),
                    _ => true,
                };
                (NUMBERING, valid)
            },
            b"pBdr" => {
                let Some(borders) = self.paragraph.borders.as_mut() else {
                    return Ok(());
                };
                let slot = match local {
                    b"top" => &mut borders.top,
                    b"left" => &mut borders.left,
                    b"bottom" => &mut borders.bottom,
                    b"right" => &mut borders.right,
                    b"between" => &mut borders.between,
                    b"bar" => &mut borders.bar,
                    _ => return Ok(()),
                };
                (BORDERS, fill_value(slot, parse_border(context, element)?))
            },
            b"tabs" => {
                if local != b"tab" {
                    return Ok(());
                }
                let Some(tabs) = self.paragraph.tabs.as_mut() else {
                    return Ok(());
                };
                let [kind, leader, position] =
                    collect(context, element, [b"val", b"leader", b"pos"])?;
                let tab = (|| {
                    Some(TabStop {
                        kind: TabKind::from_xml(kind.as_deref()?)?,
                        leader: optional(leader.as_deref(), TabLeader::from_xml)?,
                        position: Twips(measure(position.as_deref()?, Unit::Twips, Sign::Signed)?),
                    })
                })();
                let valid = tab.is_some();
                if let Some(tab) = tab {
                    tabs.push(tab);
                }
                (TABS, valid)
            },
            _ => return Ok(()),
        };
        if !valid {
            self.invalid(index);
        }
        Ok(())
    }

    fn finish(mut self) -> Paragraph {
        if let Some(mark) = self.mark {
            self.paragraph.mark = Some(Box::new(mark.finish()));
        }
        self.paragraph
    }
}

/// The field of an on/off paragraph property.
fn toggle_slot<'p>(paragraph: &'p mut Paragraph, local: &[u8]) -> Option<&'p mut Option<bool>> {
    Some(match local {
        b"keepNext" => &mut paragraph.keep_next,
        b"keepLines" => &mut paragraph.keep_lines,
        b"pageBreakBefore" => &mut paragraph.page_break_before,
        b"widowControl" => &mut paragraph.widow_control,
        b"suppressLineNumbers" => &mut paragraph.suppress_line_numbers,
        b"suppressAutoHyphens" => &mut paragraph.suppress_auto_hyphens,
        b"kinsoku" => &mut paragraph.kinsoku,
        b"wordWrap" => &mut paragraph.word_wrap,
        b"overflowPunct" => &mut paragraph.overflow_punctuation,
        b"topLinePunct" => &mut paragraph.top_line_punctuation,
        b"autoSpaceDE" => &mut paragraph.auto_space_east_asian_latin,
        b"autoSpaceDN" => &mut paragraph.auto_space_east_asian_numbers,
        b"bidi" => &mut paragraph.bidi,
        b"adjustRightInd" => &mut paragraph.adjust_right_indent,
        b"snapToGrid" => &mut paragraph.snap_to_grid,
        b"contextualSpacing" => &mut paragraph.contextual_spacing,
        b"mirrorIndents" => &mut paragraph.mirror_indents,
        b"suppressOverlap" => &mut paragraph.suppress_overlap,
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

/// Store a parsed value unless the field is already set, as a repeated
/// child element is read from its first occurrence.
fn fill_value<T>(slot: &mut Option<T>, value: Option<T>) -> bool {
    match value {
        Some(value) => {
            if slot.is_none() {
                *slot = Some(value);
            }
            true
        },
        None => false,
    }
}

/// Fill an unset attribute slot of a repeated `w:ind`, as Word merges them.
fn fill<T>(
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

/// `CT_Spacing`, typed as [`crate::Paragraph::spacing`] types it.
fn parse_spacing(
    context: &Context<'_>,
    element: &BytesStart<'_>,
) -> Result<Option<ParagraphSpacing>> {
    let [
        before,
        before_lines,
        before_auto_spacing,
        after,
        after_lines,
        after_auto_spacing,
        line,
        line_rule,
    ] = collect(
        context,
        element,
        [
            b"before",
            b"beforeLines",
            b"beforeAutospacing",
            b"after",
            b"afterLines",
            b"afterAutospacing",
            b"line",
            b"lineRule",
        ],
    )?;
    let twips = |raw: &str| {
        measure(raw, Unit::Twips, Sign::Unsigned).and_then(|value| u64::try_from(value).ok())
    };
    let lines = |raw: &str| integer(raw).and_then(|value| i32::try_from(value).ok());
    Ok((|| {
        Some(ParagraphSpacing {
            before: optional(before.as_deref(), twips)?,
            before_lines: optional(before_lines.as_deref(), lines)?,
            before_auto_spacing: optional(before_auto_spacing.as_deref(), on_off)?,
            after: optional(after.as_deref(), twips)?,
            after_lines: optional(after_lines.as_deref(), lines)?,
            after_auto_spacing: optional(after_auto_spacing.as_deref(), on_off)?,
            line: optional(line.as_deref(), |raw| {
                measure(raw, Unit::Twips, Sign::Signed).and_then(|value| i32::try_from(value).ok())
            })?,
            line_rule: optional(line_rule.as_deref(), LineSpacingRule::from_xml)?,
        })
    })())
}

/// Read a paragraph fragment's direct properties.
pub(crate) fn parse(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Paragraph> {
    let mut reader = Reader::default();
    scan::walk(
        xml,
        inherited,
        b"p",
        &[b"pPr"],
        "paragraph",
        |context, site, local, element| match site {
            Site::Property(_) => reader.property(context, local, element),
            Site::Nested(parent) => reader.nested(context, parent, local, element),
        },
    )?;
    Ok(reader.finish())
}

impl std::fmt::Debug for Paragraph {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Paragraph").finish_non_exhaustive()
    }
}
