//! The DOCX element kinds and their spec table (record 0010 §4.2, §4.5,
//! §7.2).
//!
//! [`Kind`] names every element the adapter resolves, plus the WordprocessingML
//! elements OfficeCLI has that Litchi cannot read through this adapter yet;
//! those carry a [`GapSpec`] with the follow-up that closes them, so path
//! resolution and query compilation report `unsupported_capability` and
//! `capabilities()` lists them. [`SPEC`] holds one [`ElementSpec`] per kind,
//! at the kind's index, checked at compile time by `check_format_spec!`.
//!
//! The property tables use OfficeCLI's key names and order
//! (`WordHandler.Navigation.cs`: `ParagraphToNode`, `RunToNode`,
//! `TableToNode`, `ReadRowProps`, `ReadCellProps`, `SdtBlockToNode`,
//! `HyperlinkToNode`, `GetRootNode`), restricted to the keys the leaf's typed
//! direct properties answer (L-DOCX-R4).

use crate::error::Code;
use crate::node::spec::{ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec};
use crate::path::navigate::ElementKind;
use crate::path::vocabulary::{self, Vocabulary};
use crate::registry::Constraint;
use crate::value::{ColorContext, LengthContext, LengthSpec, LengthStyle, Tokens, ValueType};

use super::setters::Setter;

/// The element kinds of a DOCX document.
///
/// Most kinds are read by this adapter. The others exist in
/// WordprocessingML and OfficeCLI but are capability gaps here: resolving
/// one is `unsupported_capability` with its follow-up.
///
/// ```
/// use litchi_ops::format::docx::Kind;
///
/// assert_eq!(Kind::Paragraph.type_name(), "paragraph");
/// assert!(Kind::Theme.is_gap());
/// assert!(!Kind::Header.is_gap());
/// assert!(!Kind::Run.is_gap());
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Kind {
    /// The document root, `/` (type `document`).
    Document,
    /// The main story's body, `/body`.
    Body,
    /// A paragraph, `p[N]`: a `w:p` that is not a display equation.
    Paragraph,
    /// A run, `r[N]`: one of the paragraph's addressable runs.
    Run,
    /// A direct `w:hyperlink` child of a paragraph, `hyperlink[N]`.
    Hyperlink,
    /// A table, `tbl[N]`.
    Table,
    /// A table row, `tr[N]`.
    Row,
    /// A table cell, `tc[N]`.
    Cell,
    /// A block content control, `sdt[N]`.
    Sdt,
    /// Display math, `oMathPara[N]`: an `m:oMathPara`, bare or held in a
    /// `w:p`.
    Equation,
    /// A body-level section break (`w:sectPr`), `section[N]`.
    Section,
    /// A header part (record 0319).
    Header,
    /// A footer part (record 0319).
    Footer,
    /// The footnotes part (record 0319).
    Footnotes,
    /// A footnote (record 0319).
    Footnote,
    /// The endnotes part (record 0319).
    Endnotes,
    /// An endnote (record 0319).
    Endnote,
    /// The comments part (record 0251c).
    Comments,
    /// A comment (record 0251c).
    Comment,
    /// The styles part (record 0319).
    Styles,
    /// One style (record 0319).
    Style,
    /// The numbering part (record 0319).
    Numbering,
    /// An abstract numbering definition (record 0319).
    AbstractNum,
    /// A numbering instance (record 0319).
    Num,
    /// A numbering level (record 0319).
    Level,
    /// The settings part (gap).
    Settings,
    /// The theme part (gap).
    Theme,
    /// A bookmark (record 0319).
    Bookmark,
    /// A field (record 0319).
    Field,
    /// A legacy form field (gap).
    FormField,
    /// An inline picture (record 0251d).
    Picture,
    /// A chart of a paragraph (record 0283).
    Chart,
    /// An embedded OLE object (gap).
    Ole,
    /// A table of contents (record 0319).
    Toc,
    /// A tracked revision (gap).
    Revision,
    /// The watermark (record 0319).
    Watermark,
    /// A text box (record 0319).
    TextBox,
    /// A drawing shape (gap).
    Shape,
    /// A run break (record 0319).
    Break,
    /// A paragraph tab stop (record 0319).
    Tab,
    /// A SmartArt graphic, addressed document-wide (record 0295).
    SmartArt,
    /// A series of a chart (record 0335).
    Series,
    /// An axis of a chart (record 0335).
    Axis,
    /// A content node of a SmartArt graphic (record 0335).
    Node,
    /// A grid column of a table, `col[N]` (record 0345): a `w:gridCol` of
    /// `w:tblGrid` and the cells at its slot.
    Column,
}

impl Kind {
    /// The kind's spec entry.
    pub(crate) fn spec(self) -> &'static ElementSpec<Self> {
        &SPEC.elements[self as usize]
    }

    /// The OfficeCLI `type` of elements of this kind (`paragraph`, `cell`).
    #[must_use]
    pub fn type_name(self) -> &'static str {
        self.spec().name
    }

    /// Whether the format has the kind but this adapter cannot read it yet:
    /// resolving it is `unsupported_capability`.
    #[must_use]
    pub fn is_gap(self) -> bool {
        self.spec().gap.is_some()
    }

    /// The leaf block kind of a kind that is a direct child of a block
    /// container (the body, a cell, a content control).
    pub(crate) const fn block_kind(self) -> Option<litchi_docx::BlockKind> {
        Some(match self {
            Self::Paragraph => litchi_docx::BlockKind::Paragraph,
            Self::Table => litchi_docx::BlockKind::Table,
            Self::Sdt => litchi_docx::BlockKind::ContentControl,
            Self::Equation => litchi_docx::BlockKind::Equation,
            Self::Section => litchi_docx::BlockKind::SectionProperties,
            _ => return None,
        })
    }

    /// The kind of a leaf block, or `None` for the blocks the adapter does
    /// not address (`w:altChunk` anchors and unknown children).
    pub(crate) fn of_block(kind: litchi_docx::BlockKind) -> Option<Self> {
        match kind {
            litchi_docx::BlockKind::Paragraph => Some(Self::Paragraph),
            litchi_docx::BlockKind::Table => Some(Self::Table),
            litchi_docx::BlockKind::ContentControl => Some(Self::Sdt),
            litchi_docx::BlockKind::Equation => Some(Self::Equation),
            litchi_docx::BlockKind::SectionProperties => Some(Self::Section),
            _ => None,
        }
    }

    /// The name a path step uses for the kind: the vocabulary's canonical
    /// name, except display equations, which OfficeCLI writes `oMathPara`
    /// (an alias of `equation`).
    pub(crate) fn step_name(self) -> &'static str {
        match self {
            Self::Equation => "oMathPara",
            _ => self.name(),
        }
    }

    /// Whether a path step names the kind without an index (`/body`).
    pub(crate) const fn is_singleton(self) -> bool {
        matches!(
            self,
            Self::Body
                | Self::Comments
                | Self::Styles
                | Self::Numbering
                | Self::Footnotes
                | Self::Endnotes
                | Self::Watermark
        )
    }
}

/// Declares the canonical vocabulary name of every kind.
macro_rules! names {
    ($($kind:ident => $canonical:literal,)+) => {
        impl ElementKind for Kind {
            const VOCABULARY: &'static Vocabulary = &vocabulary::WORD;
            const ROOT: Self = Self::Document;

            fn from_name(name: &str) -> Option<Self> {
                match name {
                    $($canonical => Some(Self::$kind),)+
                    _ => None,
                }
            }

            fn name(self) -> &'static str {
                match self {
                    $(Self::$kind => $canonical,)+
                }
            }

            fn ordinal(self) -> u8 {
                self as u8
            }
        }
    };
}

names! {
    Document => "document",
    Body => "body",
    Paragraph => "p",
    Run => "r",
    Hyperlink => "hyperlink",
    Table => "tbl",
    Row => "tr",
    Cell => "tc",
    Sdt => "sdt",
    Equation => "equation",
    Section => "section",
    Header => "header",
    Footer => "footer",
    Footnotes => "footnotes",
    Footnote => "footnote",
    Endnotes => "endnotes",
    Endnote => "endnote",
    Comments => "comments",
    Comment => "comment",
    Styles => "styles",
    Style => "style",
    Numbering => "numbering",
    AbstractNum => "abstractNum",
    Num => "num",
    Level => "level",
    Settings => "settings",
    Theme => "theme",
    Bookmark => "bookmark",
    Field => "field",
    FormField => "formfield",
    Picture => "picture",
    Chart => "chart",
    Ole => "ole",
    Toc => "toc",
    Revision => "revision",
    Watermark => "watermark",
    TextBox => "textbox",
    Shape => "shape",
    Break => "break",
    Tab => "tab",
    SmartArt => "smartart",
    Series => "series",
    Axis => "axis",
    Node => "node",
    Column => "col",
}

// ---------------------------------------------------------------------------
// Closed vocabularies
// ---------------------------------------------------------------------------

/// Paragraph `align`: `ST_Jc`, with `both` read back as `justify` as
/// OfficeCLI does.
pub(crate) static ALIGN: Tokens = Tokens::new(
    "paragraph alignment",
    &[
        "left",
        "center",
        "right",
        "justify",
        "start",
        "end",
        "distribute",
        "mediumKashida",
        "highKashida",
        "lowKashida",
        "thaiDistribute",
        "numTab",
    ],
    &[("both", 3)],
);

/// `lineRule`: `ST_LineSpacingRule`.
pub(crate) static LINE_RULE: Tokens = Tokens::new("line rule", &["auto", "exact", "atLeast"], &[]);

/// `direction`: `rtl` or `ltr`, from `w:bidi` (paragraphs) or `w:rtl` (runs).
pub(crate) static DIRECTION: Tokens = Tokens::new("direction", &["ltr", "rtl"], &[]);

/// Run `underline`: `ST_Underline`.
pub(crate) static UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "none",
        "single",
        "words",
        "double",
        "thick",
        "dotted",
        "dottedHeavy",
        "dash",
        "dashedHeavy",
        "dashLong",
        "dashLongHeavy",
        "dotDash",
        "dashDotHeavy",
        "dotDotDash",
        "dashDotDotHeavy",
        "wave",
        "wavyHeavy",
        "wavyDouble",
    ],
    &[],
);

/// Run `highlight`: `ST_HighlightColor`.
pub(crate) static HIGHLIGHT: Tokens = Tokens::new(
    "highlight",
    &[
        "black",
        "blue",
        "cyan",
        "green",
        "magenta",
        "red",
        "yellow",
        "white",
        "darkBlue",
        "darkCyan",
        "darkGreen",
        "darkMagenta",
        "darkRed",
        "darkYellow",
        "darkGray",
        "lightGray",
        "none",
    ],
    &[],
);

/// `font.hint`: `w:rFonts/@w:hint`.
pub(crate) static FONT_HINT: Tokens = Tokens::new("font hint", &["default", "eastAsia", "cs"], &[]);

/// `font.*Theme`: `ST_Theme`.
pub(crate) static THEME_FONT: Tokens = Tokens::new(
    "theme font",
    &[
        "majorEastAsia",
        "majorBidi",
        "majorAscii",
        "majorHAnsi",
        "minorEastAsia",
        "minorBidi",
        "minorAscii",
        "minorHAnsi",
    ],
    &[],
);

/// Table `align` and row `rowAlign`: `ST_JcTable`.
pub(crate) static TABLE_ALIGN: Tokens = Tokens::new(
    "table alignment",
    &["center", "end", "left", "right", "start"],
    &[],
);

/// Table `layout`, read back as OfficeCLI writes it (`autofit`, not `auto`).
pub(crate) static TABLE_LAYOUT: Tokens = Tokens::new("table layout", &["fixed", "autofit"], &[]);

/// Row `height.rule`: `ST_HeightRule`.
pub(crate) static HEIGHT_RULE: Tokens =
    Tokens::new("height rule", &["auto", "exact", "atLeast"], &[]);

/// Cell `valign`: `ST_VerticalJc`.
pub(crate) static VALIGN: Tokens = Tokens::new(
    "cell vertical alignment",
    &["top", "center", "both", "bottom"],
    &[],
);

/// Cell `vmerge` and `hmerge`: `ST_Merge`.
pub(crate) static MERGE: Tokens = Tokens::new("merge", &["restart", "continue"], &[]);

/// Cell `textDirection`: `ST_TextDirection`.
pub(crate) static TEXT_DIRECTION: Tokens = Tokens::new(
    "text direction",
    &[
        "tb", "rl", "lr", "tbV", "rlV", "lrV", "btLr", "lrTb", "lrTbV", "tbLrV", "tbRl", "tbRlV",
    ],
    &[],
);

// ---------------------------------------------------------------------------
// Length types
// ---------------------------------------------------------------------------

/// Paragraph spacing: non-negative, printed in points (`12pt`), as
/// OfficeCLI's `FormatWordSpacingNonNegative` prints it.
pub(crate) const SPACING: LengthSpec = LengthSpec::new(LengthContext::Extent, LengthStyle::Points);

/// Indents, baseline offsets and character spacing: signed, printed in
/// points (`-18pt`), as `FormatWordSpacing` prints them.
pub(crate) const SIGNED_POINTS: LengthSpec =
    LengthSpec::new(LengthContext::Emu, LengthStyle::Points);

// ---------------------------------------------------------------------------
// Property tables
// ---------------------------------------------------------------------------

/// One writable DOCX declaration.  The same typed setter can construct a
/// property on a newly added node and update it on an existing node, so the
/// registry exposes both operations.  The read projection is still emitted
/// in declaration order, so adding a write-only key at the end of a table
/// does not change any existing node property index.
const fn writable(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
    clear: bool,
) -> crate::registry::Entry<Setter> {
    crate::registry::Entry {
        spec: PropSpec::new(name, ty),
        access: crate::registry::Access::SetAndAdd(setter),
        constraint,
        clear,
    }
}

/// A property accepted only while creating a paragraph's initial run.
///
/// These rows deliberately live after the paragraph's established read and
/// write projection.  They are consumed by the structural `add` handler when
/// it builds the first `NewRun`; an existing paragraph never exposes them as
/// direct paragraph properties.
const fn add_only_run(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    crate::registry::Entry {
        spec,
        access: crate::registry::Access::Add(setter),
        constraint,
        clear: true,
    }
}

/// A read property that also seeds a newly added element (record 0251a).
const fn read_add(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    crate::registry::Entry {
        spec: PropSpec::new(name, ty),
        access: crate::registry::Access::Add(setter),
        constraint,
        clear: false,
    }
}

/// A property accepted only while adding an element.
const fn add_only(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    crate::registry::Entry {
        spec,
        access: crate::registry::Access::Add(setter),
        constraint,
        clear: false,
    }
}

/// A property that `get` reads and `set` writes, and `add` does not seed.
const fn read_set(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    crate::registry::Entry {
        spec: PropSpec::new(name, ty),
        access: crate::registry::Access::Set(setter),
        constraint,
        clear: false,
    }
}

/// A property that `get` reads, `set` writes and `add` seeds, with no clear
/// (record 0342: `listStyle` and `start` on a new paragraph).
const fn read_set_add(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    crate::registry::Entry {
        spec: PropSpec::new(name, ty),
        access: crate::registry::Access::SetAndAdd(setter),
        constraint,
        clear: false,
    }
}

/// A write-only property that `set` writes and `add` seeds (record 0342:
/// table borders); `null` removes it.
const fn write_set_add(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    crate::registry::Entry {
        spec,
        access: crate::registry::Access::SetAndAdd(setter),
        constraint,
        clear: true,
    }
}

/// A write-only property that only `set` writes (record 0342: the row
/// height rules, which read back as `height` and `height.rule`).
const fn write_set(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
    constraint: Constraint,
) -> crate::registry::Entry<Setter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    crate::registry::Entry {
        spec,
        access: crate::registry::Access::Set(setter),
        constraint,
        clear: true,
    }
}

/// One write-only property declaration.  Text has a dedicated write handler
/// but is intentionally absent from the read model; the appended registry
/// row keeps all pre-existing paragraph/run read indices stable.
const fn write_only(
    name: &'static str,
    ty: ValueType,
    setter: Setter,
) -> crate::registry::Entry<Setter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    crate::registry::Entry {
        spec,
        access: crate::registry::Access::SetAndAdd(setter),
        constraint: Constraint::None,
        clear: true,
    }
}

// `GetRootNode`: the core document properties.
crate::registry::registry! {
    pub(crate) DOCUMENT_PROPS, DOCUMENT_REGISTRY: () {
    crate::registry::Entry::read(PropSpec::new("title", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("author", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("subject", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("keywords", ValueType::List)),
    crate::registry::Entry::read(PropSpec::new("description", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("category", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("lastModifiedBy", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("revisionNumber", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("created", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("modified", ValueType::Text)),
    // `ExtendedPropertiesHandler`: the application properties (record 0342).
    crate::registry::Entry::read(PropSpec::new("extended.template", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("extended.manager", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("extended.company", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("extended.application", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("extended.applicationVersion", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("extended.pages", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("extended.words", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("extended.characters", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("extended.lines", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("extended.paragraphs", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("extended.totalTime", ValueType::Int)),
    }
}

/// The root's capability table: the core-property keys and the settable
/// application keys are settable, through the metadata writer on `/` and
/// `/metadata` (records 0338 and 0342). Same order as [`DOCUMENT_PROPS`],
/// so reads index either.
pub(crate) static DOCUMENT_ROOT_PROPS: [PropSpec; DOCUMENT_PROPS.len()] =
    crate::format::metadata::root_capability(DOCUMENT_PROPS);

// `ParagraphToNode`, in its order. `numId` and `numLevel` are strings
// there, `outlineLvl` an int.
crate::registry::registry! {
    pub(crate) PARAGRAPH_PROPS, PARAGRAPH_REGISTRY: Setter {
    crate::registry::Entry::read(PropSpec::new("paraId", ValueType::Text)),
    writable("style", ValueType::Text, Setter::ParagraphStyle, Constraint::NotEmpty, true),
    // Record 0319: `styleId` writes the id; `styleName` reads the style's
    // display name and writes a style named by id or display name.
    writable("styleId", ValueType::Text, Setter::ParagraphStyle, Constraint::NotEmpty, true),
    writable("styleName", ValueType::Text, Setter::ParagraphStyle, Constraint::NotEmpty, true),
    writable("align", ValueType::Token(&ALIGN), Setter::ParagraphAlign, Constraint::None, true),
    writable("spaceBefore", ValueType::Length(SPACING), Setter::ParagraphSpaceBefore, Constraint::None, true),
    writable("spaceAfter", ValueType::Length(SPACING), Setter::ParagraphSpaceAfter, Constraint::None, true),
    writable("lineSpacing", ValueType::Text, Setter::ParagraphLineSpacing, Constraint::NotEmpty, true),
    writable("lineRule", ValueType::Token(&LINE_RULE), Setter::ParagraphLineRule, Constraint::None, true),
    writable("spaceBeforeLines", ValueType::Int, Setter::ParagraphSpaceBeforeLines, Constraint::None, true),
    writable("spaceAfterLines", ValueType::Int, Setter::ParagraphSpaceAfterLines, Constraint::None, true),
    writable("spaceBeforeAuto", ValueType::Bool, Setter::ParagraphSpaceBeforeAuto, Constraint::None, true),
    writable("spaceAfterAuto", ValueType::Bool, Setter::ParagraphSpaceAfterAuto, Constraint::None, true),
    writable("firstLineIndent", ValueType::Length(SIGNED_POINTS), Setter::ParagraphFirstLineIndent, Constraint::None, true),
    writable("hangingIndent", ValueType::Length(SIGNED_POINTS), Setter::ParagraphHangingIndent, Constraint::None, true),
    writable("indent", ValueType::Length(SIGNED_POINTS), Setter::ParagraphIndent, Constraint::None, true),
    writable("rightIndent", ValueType::Length(SIGNED_POINTS), Setter::ParagraphRightIndent, Constraint::None, true),
    writable("firstLineChars", ValueType::Int, Setter::ParagraphFirstLineChars, Constraint::None, true),
    writable("hangingChars", ValueType::Int, Setter::ParagraphHangingChars, Constraint::None, true),
    writable("leftChars", ValueType::Int, Setter::ParagraphLeftChars, Constraint::None, true),
    writable("rightChars", ValueType::Int, Setter::ParagraphRightChars, Constraint::None, true),
    writable("keepNext", ValueType::Bool, Setter::ParagraphKeepNext, Constraint::None, true),
    writable("keepLines", ValueType::Bool, Setter::ParagraphKeepLines, Constraint::None, true),
    writable("pageBreakBefore", ValueType::Bool, Setter::ParagraphPageBreakBefore, Constraint::None, true),
    writable("widowControl", ValueType::Bool, Setter::ParagraphWidowControl, Constraint::None, true),
    writable("direction", ValueType::Token(&DIRECTION), Setter::ParagraphDirection, Constraint::None, true),
    writable("contextualSpacing", ValueType::Bool, Setter::ParagraphContextualSpacing, Constraint::None, true),
    writable("numId", ValueType::Text, Setter::ParagraphNumId, Constraint::NotEmpty, true),
    writable("numLevel", ValueType::Text, Setter::ParagraphNumLevel, Constraint::NotEmpty, true),
    // Record 0319: the numbering part answers the paragraph's list format,
    // kind and start (record 0251e writes the kind and start).
    crate::registry::Entry::read(PropSpec::new("numFmt", ValueType::Text)),
    // Record 0342: `add` creates or continues the list too.
    read_set_add("listStyle", ValueType::Text, Setter::ParagraphListStyle, Constraint::NotEmpty),
    read_set_add("start", ValueType::Int, Setter::ParagraphListStart, Constraint::Int { min: 0, max: 32_767 }),
    writable("outlineLvl", ValueType::Int, Setter::ParagraphOutlineLevel, Constraint::Int { min: 0, max: 9 }, true),
    write_only("text", ValueType::Text, Setter::ParagraphText),
    // A paragraph may carry one initial run.  Its direct run formatting is
    // add-only metadata: it is accepted by `add paragraph`, but is not a
    // paragraph property for `get`, `query` or `set`.  Keep paragraph-owned
    // `direction`, `style` and `text` authoritative above.
    add_only_run("font.latin", ValueType::Text, Setter::RunFontLatin, Constraint::NotEmpty),
    add_only_run("font.ascii", ValueType::Text, Setter::RunFontAscii, Constraint::NotEmpty),
    add_only_run("font.hAnsi", ValueType::Text, Setter::RunFontHAnsi, Constraint::NotEmpty),
    add_only_run("font.ea", ValueType::Text, Setter::RunFontEastAsia, Constraint::NotEmpty),
    add_only_run("font.asciiTheme", ValueType::Token(&THEME_FONT), Setter::RunFontAsciiTheme, Constraint::None),
    add_only_run("font.hAnsiTheme", ValueType::Token(&THEME_FONT), Setter::RunFontHAnsiTheme, Constraint::None),
    add_only_run("font.eaTheme", ValueType::Token(&THEME_FONT), Setter::RunFontEastAsiaTheme, Constraint::None),
    add_only_run("font.csTheme", ValueType::Token(&THEME_FONT), Setter::RunFontComplexScriptTheme, Constraint::None),
    add_only_run("font.hint", ValueType::Token(&FONT_HINT), Setter::RunFontHint, Constraint::None),
    add_only_run("lang.latin", ValueType::Text, Setter::RunLanguageLatin, Constraint::NotEmpty),
    add_only_run("lang.ea", ValueType::Text, Setter::RunLanguageEastAsia, Constraint::NotEmpty),
    add_only_run("lang.cs", ValueType::Text, Setter::RunLanguageComplexScript, Constraint::NotEmpty),
    add_only_run("size", ValueType::Length(LengthSpec::FONT_SIZE), Setter::RunSize, Constraint::None),
    add_only_run("bold", ValueType::Bool, Setter::RunBold, Constraint::None),
    add_only_run("italic", ValueType::Bool, Setter::RunItalic, Constraint::None),
    add_only_run("font.cs", ValueType::Text, Setter::RunFontComplexScript, Constraint::NotEmpty),
    add_only_run("size.cs", ValueType::Length(LengthSpec::FONT_SIZE), Setter::RunSizeComplexScript, Constraint::None),
    add_only_run("bold.cs", ValueType::Bool, Setter::RunBoldComplexScript, Constraint::None),
    add_only_run("italic.cs", ValueType::Bool, Setter::RunItalicComplexScript, Constraint::None),
    add_only_run("color", ValueType::Color(ColorContext::Word), Setter::RunColor, Constraint::None),
    add_only_run("underline", ValueType::Token(&UNDERLINE), Setter::RunUnderline, Constraint::None),
    add_only_run("underline.color", ValueType::Color(ColorContext::Word), Setter::RunUnderlineColor, Constraint::None),
    add_only_run("strike", ValueType::Bool, Setter::RunStrike, Constraint::None),
    add_only_run("highlight", ValueType::Token(&HIGHLIGHT), Setter::RunHighlight, Constraint::None),
    add_only_run("caps", ValueType::Bool, Setter::RunCaps, Constraint::None),
    add_only_run("smallcaps", ValueType::Bool, Setter::RunSmallCaps, Constraint::None),
    add_only_run("dstrike", ValueType::Bool, Setter::RunDoubleStrike, Constraint::None),
    add_only_run("vanish", ValueType::Bool, Setter::RunVanish, Constraint::None),
    add_only_run("outline", ValueType::Bool, Setter::RunOutline, Constraint::None),
    add_only_run("shadow", ValueType::Bool, Setter::RunShadow, Constraint::None),
    add_only_run("emboss", ValueType::Bool, Setter::RunEmboss, Constraint::None),
    add_only_run("imprint", ValueType::Bool, Setter::RunImprint, Constraint::None),
    add_only_run("noproof", ValueType::Bool, Setter::RunNoProof, Constraint::None),
    add_only_run("superscript", ValueType::Bool, Setter::RunSuperscript, Constraint::None),
    add_only_run("subscript", ValueType::Bool, Setter::RunSubscript, Constraint::None),
    add_only_run("position", ValueType::Length(SIGNED_POINTS), Setter::RunPosition, Constraint::None),
    add_only_run("charSpacing", ValueType::Length(SIGNED_POINTS), Setter::RunCharacterSpacing, Constraint::None),
    add_only_run("rStyle", ValueType::Text, Setter::RunStyle, Constraint::NotEmpty),
    }
}

// `RunToNode` (with `ReadComplexScriptRunFormatting`), in its order.
crate::registry::registry! {
    pub(crate) RUN_PROPS, RUN_REGISTRY: Setter {
    writable("font.latin", ValueType::Text, Setter::RunFontLatin, Constraint::NotEmpty, true),
    writable("font.ascii", ValueType::Text, Setter::RunFontAscii, Constraint::NotEmpty, true),
    writable("font.hAnsi", ValueType::Text, Setter::RunFontHAnsi, Constraint::NotEmpty, true),
    writable("font.ea", ValueType::Text, Setter::RunFontEastAsia, Constraint::NotEmpty, true),
    writable("font.asciiTheme", ValueType::Token(&THEME_FONT), Setter::RunFontAsciiTheme, Constraint::None, true),
    writable("font.hAnsiTheme", ValueType::Token(&THEME_FONT), Setter::RunFontHAnsiTheme, Constraint::None, true),
    writable("font.eaTheme", ValueType::Token(&THEME_FONT), Setter::RunFontEastAsiaTheme, Constraint::None, true),
    writable("font.csTheme", ValueType::Token(&THEME_FONT), Setter::RunFontComplexScriptTheme, Constraint::None, true),
    writable("font.hint", ValueType::Token(&FONT_HINT), Setter::RunFontHint, Constraint::None, true),
    writable("lang.latin", ValueType::Text, Setter::RunLanguageLatin, Constraint::NotEmpty, true),
    writable("lang.ea", ValueType::Text, Setter::RunLanguageEastAsia, Constraint::NotEmpty, true),
    writable("lang.cs", ValueType::Text, Setter::RunLanguageComplexScript, Constraint::NotEmpty, true),
    writable("size", ValueType::Length(LengthSpec::FONT_SIZE), Setter::RunSize, Constraint::None, true),
    writable("bold", ValueType::Bool, Setter::RunBold, Constraint::None, true),
    writable("italic", ValueType::Bool, Setter::RunItalic, Constraint::None, true),
    writable("font.cs", ValueType::Text, Setter::RunFontComplexScript, Constraint::NotEmpty, true),
    writable("size.cs", ValueType::Length(LengthSpec::FONT_SIZE), Setter::RunSizeComplexScript, Constraint::None, true),
    writable("bold.cs", ValueType::Bool, Setter::RunBoldComplexScript, Constraint::None, true),
    writable("italic.cs", ValueType::Bool, Setter::RunItalicComplexScript, Constraint::None, true),
    writable("color", ValueType::Color(ColorContext::Word), Setter::RunColor, Constraint::None, true),
    writable("underline", ValueType::Token(&UNDERLINE), Setter::RunUnderline, Constraint::None, true),
    writable("underline.color", ValueType::Color(ColorContext::Word), Setter::RunUnderlineColor, Constraint::None, true),
    writable("strike", ValueType::Bool, Setter::RunStrike, Constraint::None, true),
    writable("highlight", ValueType::Token(&HIGHLIGHT), Setter::RunHighlight, Constraint::None, true),
    writable("caps", ValueType::Bool, Setter::RunCaps, Constraint::None, true),
    writable("smallcaps", ValueType::Bool, Setter::RunSmallCaps, Constraint::None, true),
    writable("dstrike", ValueType::Bool, Setter::RunDoubleStrike, Constraint::None, true),
    writable("vanish", ValueType::Bool, Setter::RunVanish, Constraint::None, true),
    writable("outline", ValueType::Bool, Setter::RunOutline, Constraint::None, true),
    writable("shadow", ValueType::Bool, Setter::RunShadow, Constraint::None, true),
    writable("emboss", ValueType::Bool, Setter::RunEmboss, Constraint::None, true),
    writable("imprint", ValueType::Bool, Setter::RunImprint, Constraint::None, true),
    writable("noproof", ValueType::Bool, Setter::RunNoProof, Constraint::None, true),
    writable("direction", ValueType::Token(&DIRECTION), Setter::RunDirection, Constraint::None, true),
    writable("superscript", ValueType::Bool, Setter::RunSuperscript, Constraint::None, true),
    writable("subscript", ValueType::Bool, Setter::RunSubscript, Constraint::None, true),
    writable("position", ValueType::Length(SIGNED_POINTS), Setter::RunPosition, Constraint::None, true),
    writable("charSpacing", ValueType::Length(SIGNED_POINTS), Setter::RunCharacterSpacing, Constraint::None, true),
    writable("rStyle", ValueType::Text, Setter::RunStyle, Constraint::NotEmpty, true),
    write_only("text", ValueType::Text, Setter::RunText),
    }
}

// `HyperlinkToNode`, in its order: the link's attributes, then the
// formatting of its first run that holds text.
crate::registry::registry! {
    pub(crate) HYPERLINK_PROPS, HYPERLINK_REGISTRY: Setter {
    crate::registry::Entry {
        spec: PropSpec::new("url", ValueType::Text),
        access: crate::registry::Access::SetAndAdd(Setter::HyperlinkUrl),
        constraint: Constraint::NotEmpty,
        clear: false,
    },
    read_add("anchor", ValueType::Text, Setter::HyperlinkAnchor, Constraint::NotEmpty),
    crate::registry::Entry::read(PropSpec::new("docLocation", ValueType::Text)),
    // Record 0345: the screen tip is written by add and set; `null` removes it.
    writable("tooltip", ValueType::Text, Setter::HyperlinkTooltip, Constraint::None, true),
    crate::registry::Entry::read(PropSpec::new("tgtFrame", ValueType::Text)),
    read_add("font", ValueType::Text, Setter::RunFontLatin, Constraint::NotEmpty),
    crate::registry::Entry::read(PropSpec::new("font.cs", ValueType::Text)),
    read_add("size", ValueType::Length(LengthSpec::FONT_SIZE), Setter::RunSize, Constraint::None),
    read_add("bold", ValueType::Bool, Setter::RunBold, Constraint::None),
    read_add("italic", ValueType::Bool, Setter::RunItalic, Constraint::None),
    read_add("color", ValueType::Color(ColorContext::Word), Setter::RunColor, Constraint::None),
    read_add("underline", ValueType::Token(&UNDERLINE), Setter::RunUnderline, Constraint::None),
    crate::registry::Entry::read(PropSpec::new("underline.color", ValueType::Color(ColorContext::Word))),
    crate::registry::Entry::read(PropSpec::new("strike", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("highlight", ValueType::Token(&HIGHLIGHT))),
    read_add("rStyle", ValueType::Text, Setter::RunStyle, Constraint::NotEmpty),
    write_only("text", ValueType::Text, Setter::HyperlinkText),
    }
}

// `TableToNode`, in its order. `cols` and `rows` are strings there;
// `width` is a string (`"5000"`, `"50%"`, `"auto"`) and `indent` a number
// of twips or a percentage string.
crate::registry::registry! {
    pub(crate) TABLE_PROPS, TABLE_REGISTRY: Setter {
    read_add("cols", ValueType::Text, Setter::TableColumns, Constraint::NotEmpty),
    read_add("rows", ValueType::Text, Setter::TableRows, Constraint::NotEmpty),
    // Record 0342: the style is set too.
    writable("style", ValueType::Text, Setter::TableStyle, Constraint::NotEmpty, false),
    crate::registry::Entry::read(PropSpec::new("rowBandSize", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("colBandSize", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("width", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("align", ValueType::Token(&TABLE_ALIGN))),
    crate::registry::Entry::read(PropSpec::new("indent", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("cellSpacing", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("layout", ValueType::Token(&TABLE_LAYOUT))),
    crate::registry::Entry::read(PropSpec::new("tblLook", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("direction", ValueType::Token(&DIRECTION))),
    crate::registry::Entry::read(PropSpec::new("padding.top", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.bottom", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.left", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.right", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("firstRow", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("lastRow", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("firstCol", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("lastCol", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("bandedRows", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("bandedCols", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("caption", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("description", ValueType::Text)),
    read_add("colWidths", ValueType::Text, Setter::TableColumnWidths, Constraint::NotEmpty),
    add_only("data", ValueType::Text, Setter::TableData, Constraint::None),
    // Record 0342: `ApplyTableBorders`, `STYLE;SIZE;COLOR[;SPACE]`.
    write_set_add("border.all", ValueType::Text, Setter::TableBorderAll, Constraint::NotEmpty),
    write_set_add("border.top", ValueType::Text, Setter::TableBorderTop, Constraint::NotEmpty),
    write_set_add("border.bottom", ValueType::Text, Setter::TableBorderBottom, Constraint::NotEmpty),
    write_set_add("border.left", ValueType::Text, Setter::TableBorderLeft, Constraint::NotEmpty),
    write_set_add("border.right", ValueType::Text, Setter::TableBorderRight, Constraint::NotEmpty),
    write_set_add("border.horizontal", ValueType::Text, Setter::TableBorderHorizontal, Constraint::NotEmpty),
    write_set_add("border.vertical", ValueType::Text, Setter::TableBorderVertical, Constraint::NotEmpty),
    }
}

// `ReadRowProps`, in its order. Heights and widths are strings with a
// `dxa` unit there, grid counts strings.
crate::registry::registry! {
    pub(crate) ROW_PROPS, ROW_REGISTRY: Setter {
    // Record 0342: `height`, `header`, `cantSplit` and `hidden` are set.
    read_set("height", ValueType::Text, Setter::RowHeight, Constraint::NotEmpty),
    crate::registry::Entry::read(PropSpec::new("height.rule", ValueType::Token(&HEIGHT_RULE))),
    read_set("header", ValueType::Bool, Setter::RowHeader, Constraint::None),
    read_set("cantSplit", ValueType::Bool, Setter::RowCantSplit, Constraint::None),
    read_set("hidden", ValueType::Bool, Setter::RowHidden, Constraint::None),
    crate::registry::Entry::read(PropSpec::new("rowAlign", ValueType::Token(&TABLE_ALIGN))),
    crate::registry::Entry::read(PropSpec::new("gridBefore", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("wBefore", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("gridAfter", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("wAfter", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("cellSpacing", ValueType::Int)),
    add_only("cols", ValueType::Int, Setter::RowColumns, Constraint::Int { min: 1, max: 4096 }),
    write_set("height.exact", ValueType::Text, Setter::RowHeightExact, Constraint::NotEmpty),
    write_set("height.atLeast", ValueType::Text, Setter::RowHeightAtLeast, Constraint::NotEmpty),
    }
}

// A table grid column (record 0345), after OfficeCLI's virtual `column`
// element: `width` (twips) and `text` seed a new column; Litchi also reads
// and sets the width (the `w:gridCol`).
crate::registry::registry! {
    pub(crate) COLUMN_PROPS, COLUMN_REGISTRY: Setter {
    read_set_add("width", ValueType::Text, Setter::ColumnWidth, Constraint::NotEmpty),
    add_only("text", ValueType::Text, Setter::ColumnText, Constraint::None),
    }
}

// `ReadCellProps`, in its order. `width` is a string (`"2880dxa"`,
// `"50%"`, `"auto"`, `"nil"`) there.
crate::registry::registry! {
    pub(crate) CELL_PROPS, CELL_REGISTRY: Setter {
    // Record 0345: `width` and `text` seed a new cell (OfficeCLI's
    // `AddCell` honours only these two at add time).
    read_add("width", ValueType::Text, Setter::CellWidth, Constraint::NotEmpty),
    crate::registry::Entry::read(PropSpec::new("valign", ValueType::Token(&VALIGN))),
    crate::registry::Entry::read(PropSpec::new("vmerge", ValueType::Token(&MERGE))),
    crate::registry::Entry::read(PropSpec::new("hmerge", ValueType::Token(&MERGE))),
    crate::registry::Entry::read(PropSpec::new("colspan", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.top", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.bottom", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.left", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("padding.right", ValueType::Int)),
    crate::registry::Entry::read(PropSpec::new("textDirection", ValueType::Token(&TEXT_DIRECTION))),
    crate::registry::Entry::read(PropSpec::new("nowrap", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("hideMark", ValueType::Bool)),
    crate::registry::Entry::read(PropSpec::new("tcFitText", ValueType::Bool)),
    write_only("text", ValueType::Text, Setter::CellText),
    }
}

/// An inline picture of a paragraph (record 0251d): its extent, name,
/// alternative text and image relationship; `src` (a `data:` URI or a file
/// path) seeds a new one. Each but `relId` is also set (record 0335); a
/// `src` set replaces the image. Aliases follow OfficeCLI's picture schema.
pub(crate) const PICTURE_REGISTRY: crate::registry::Table<Setter, 6, 5> =
    crate::registry::Table::new(
        [
            writable(
                "width",
                ValueType::Length(LengthSpec::EXTENT),
                Setter::PictureWidth,
                Constraint::None,
                false,
            ),
            writable(
                "height",
                ValueType::Length(LengthSpec::EXTENT),
                Setter::PictureHeight,
                Constraint::None,
                false,
            ),
            writable(
                "alt",
                ValueType::Text,
                Setter::PictureAlt,
                Constraint::None,
                false,
            ),
            writable(
                "name",
                ValueType::Text,
                Setter::PictureName,
                Constraint::NotEmpty,
                false,
            ),
            crate::registry::Entry::read(PropSpec::new("relId", ValueType::Text)),
            {
                let mut spec = PropSpec::new("src", ValueType::Text);
                spec.access.get = false;
                crate::registry::Entry {
                    spec,
                    access: crate::registry::Access::SetAndAdd(Setter::PictureSrc),
                    constraint: Constraint::NotEmpty,
                    clear: false,
                }
            },
        ],
        [
            crate::registry::Alias {
                name: "w",
                entry: 0,
            },
            crate::registry::Alias {
                name: "h",
                entry: 1,
            },
            crate::registry::Alias {
                name: "altText",
                entry: 2,
            },
            crate::registry::Alias {
                name: "description",
                entry: 2,
            },
            crate::registry::Alias {
                name: "path",
                entry: 5,
            },
        ],
    );

/// The picture element's properties.
pub(crate) const PICTURE_PROPS: &[PropSpec] = &PICTURE_REGISTRY.readable();

// A comment of the comments part (record 0251c): its identity, author and
// text; `text` and the author's details seed a new comment and are set
// (record 0335). Record 0345, after OfficeCLI's comment schema: the
// anchor's path (`anchoredTo`), the resolved state (`done`, alias
// `resolved`; set too) and the reply parent (`parentId`, the parent's
// comment id) from `commentsExtended`, and the anchor controls `runStart`
// and `range` on add.
pub(crate) const COMMENT_REGISTRY: crate::registry::Table<Setter, 10, 1> =
    crate::registry::Table::new(
        [
            crate::registry::Entry::read(PropSpec::new("id", ValueType::Int)),
            writable(
                "author",
                ValueType::Text,
                Setter::CommentAuthor,
                Constraint::NotEmpty,
                false,
            ),
            writable(
                "initials",
                ValueType::Text,
                Setter::CommentInitials,
                Constraint::NotEmpty,
                false,
            ),
            writable(
                "date",
                ValueType::Text,
                Setter::CommentDate,
                Constraint::NotEmpty,
                false,
            ),
            writable(
                "text",
                ValueType::Text,
                Setter::CommentText,
                Constraint::None,
                false,
            ),
            crate::registry::Entry::read(PropSpec::new("anchoredTo", ValueType::Text)),
            writable(
                "done",
                ValueType::Bool,
                Setter::CommentDone,
                Constraint::None,
                false,
            ),
            read_add(
                "parentId",
                ValueType::Int,
                Setter::CommentParent,
                Constraint::Int {
                    min: 0,
                    max: u32::MAX as i64,
                },
            ),
            add_only(
                "runStart",
                ValueType::Int,
                Setter::CommentRunStart,
                Constraint::Int {
                    min: 0,
                    max: 1 << 20,
                },
            ),
            add_only(
                "range",
                ValueType::Bool,
                Setter::CommentRange,
                Constraint::None,
            ),
        ],
        [crate::registry::Alias {
            name: "resolved",
            entry: 6,
        }],
    );

/// The comment element's properties.
pub(crate) const COMMENT_PROPS: &[PropSpec] = &COMMENT_REGISTRY.readable();

// `SdtBlockToNode`: the content control's identity.
crate::registry::registry! {
    pub(crate) SDT_PROPS, SDT_REGISTRY: () {
    crate::registry::Entry::read(PropSpec::new("alias", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("tag", ValueType::Text)),
    crate::registry::Entry::read(PropSpec::new("sdtId", ValueType::Int)),
    }
}

/// An equation's `mode`: OfficeCLI writes `display` for block equations.
pub(crate) static EQUATION_MODE: Tokens = Tokens::new("equation mode", &["display", "inline"], &[]);

// `ElementToNode` for `oMathPara`: the equation's mode.
crate::registry::registry! {
    pub(crate) EQUATION_PROPS, EQUATION_REGISTRY: () {
    crate::registry::Entry::read(PropSpec::new("mode", ValueType::Token(&EQUATION_MODE))),
    }
}

/// The index of `key` in `props`, checked at compile time.
pub(crate) const fn index_of(props: &[PropSpec], key: &str) -> usize {
    let wanted = key.as_bytes();
    let mut index = 0;
    while index < props.len() {
        let name = props[index].name.as_bytes();
        if name.len() == wanted.len() {
            let mut byte = 0;
            while byte < name.len() && name[byte] == wanted[byte] {
                byte += 1;
            }
            if byte == name.len() {
                return index;
            }
        }
        index += 1;
    }
    panic!("the key is not in the property table");
}

// ---------------------------------------------------------------------------
// Gaps
// ---------------------------------------------------------------------------

/// Declares one element gap per listed kind.
macro_rules! gaps {
    ($($name:ident: $element:literal, $follow_up:literal, $reason:literal;)+) => {
        $(
            static $name: GapSpec = GapSpec {
                subject: Subject::Element($element),
                code: Code::UnsupportedCapability,
                follow_up: $follow_up,
                reason: $reason,
            };
        )+
    };
}

gaps! {
    SETTINGS_GAP: "settings", "L-DOCX-R6 (R4x)", "the settings part has no node adapter yet";
    THEME_GAP: "theme", "L-DOCX-R6 (R4x)", "the theme part has no node adapter yet";
    FORM_FIELD_GAP: "formfield", "R4x", "legacy form fields have no located reader yet";
    OLE_GAP: "ole", "R4x", "embedded objects have no located reader yet";
    REVISION_GAP: "revision", "R4x", "tracked revisions have no located reader yet";
    SHAPE_GAP: "shape", "R4x", "drawing shapes have no located reader yet";
}

/// A gap for an OfficeCLI element the table has no kind for (record 0339):
/// naming it is `unsupported_capability`.
const fn unlisted_gap(
    element: &'static str,
    follow_up: &'static str,
    reason: &'static str,
) -> GapSpec {
    GapSpec {
        subject: Subject::Element(element),
        code: Code::UnsupportedCapability,
        follow_up,
        reason,
    }
}

/// An element spec with the fields most kinds share.
const fn element(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    children: &'static [Kind],
    paths: &'static [&'static str],
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children,
        identities: &[],
        paths,
        props,
        gap: None,
    }
}

/// An element the leaf cannot answer yet.
const fn gap(
    kind: Kind,
    name: &'static str,
    aliases: &'static [&'static str],
    parents: &'static [Kind],
    paths: &'static [&'static str],
    gap: &'static GapSpec,
) -> ElementSpec<Kind> {
    ElementSpec {
        kind,
        name,
        aliases,
        parents,
        children: &[],
        identities: &[],
        paths,
        props: &[],
        gap: Some(gap),
    }
}

const fn identities(
    mut spec: ElementSpec<Kind>,
    identities: &'static [&'static str],
) -> ElementSpec<Kind> {
    spec.identities = identities;
    spec
}

/// The containers whose direct block children are paragraphs, tables,
/// content controls and equations.
const CONTAINERS: &[Kind] = &[Kind::Body, Kind::Cell, Kind::Sdt];

/// The block children of a cell or content control (a section break occurs
/// only in the body).
const BLOCKS: &[Kind] = &[Kind::Paragraph, Kind::Table, Kind::Sdt, Kind::Equation];

/// The DOCX spec table.
pub(crate) static SPEC: FormatSpec<Kind> = {
    use Kind as K;
    FormatSpec {
        root: K::Document,
        elements: &[
            element(
                K::Document,
                "document",
                &["document"],
                &[],
                &[
                    K::Body,
                    K::Styles,
                    K::Header,
                    K::Footer,
                    K::Numbering,
                    K::Footnotes,
                    K::Endnotes,
                    K::Comments,
                    K::SmartArt,
                    K::Bookmark,
                    K::Field,
                    K::Toc,
                    K::Watermark,
                ],
                &["/"],
                &DOCUMENT_ROOT_PROPS,
            ),
            element(
                K::Body,
                "body",
                &["body"],
                &[K::Document],
                &[
                    K::Paragraph,
                    K::Table,
                    K::Sdt,
                    K::Equation,
                    K::Section,
                    K::TextBox,
                ],
                &["/body"],
                &[],
            ),
            identities(
                element(
                    K::Paragraph,
                    "paragraph",
                    &["p", "paragraph"],
                    CONTAINERS,
                    &[K::Run, K::Hyperlink, K::Picture, K::Chart, K::Tab, K::Break],
                    &[
                        "/body/p[N]",
                        "/body/p[@paraId=X]",
                        "/body/tbl[N]/tr[M]/tc[K]/p[J]",
                        "/body/sdt[N]/p[M]",
                        "/header[N]/p[M]",
                        "/footnote[N]/p[M]",
                    ],
                    PARAGRAPH_PROPS,
                ),
                &["paraId", "textId"],
            ),
            element(
                K::Run,
                "run",
                &["r", "run"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/r[M]"],
                RUN_PROPS,
            ),
            element(
                K::Hyperlink,
                "hyperlink",
                &["hyperlink"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/hyperlink[M]"],
                HYPERLINK_PROPS,
            ),
            element(
                K::Table,
                "table",
                &["table", "tbl"],
                CONTAINERS,
                &[K::Row],
                &["/body/tbl[N]", "/body/tbl[N]/tr[M]/tc[K]/tbl[J]"],
                TABLE_PROPS,
            ),
            element(
                K::Row,
                "row",
                &["row", "tr"],
                &[K::Table],
                &[K::Cell],
                &["/body/tbl[N]/tr[M]"],
                ROW_PROPS,
            ),
            element(
                K::Cell,
                "cell",
                &["cell", "tc"],
                &[K::Row],
                BLOCKS,
                &["/body/tbl[N]/tr[M]/tc[K]"],
                CELL_PROPS,
            ),
            identities(
                element(
                    K::Sdt,
                    "sdt",
                    &["contentcontrol", "sdt"],
                    CONTAINERS,
                    BLOCKS,
                    &["/body/sdt[N]", "/body/sdt[@sdtId=X]"],
                    SDT_PROPS,
                ),
                &["sdtId"],
            ),
            element(
                K::Equation,
                "equation",
                &["equation", "formula", "math", "omath", "omathpara"],
                CONTAINERS,
                &[],
                &["/body/oMathPara[N]"],
                EQUATION_PROPS,
            ),
            element(
                K::Section,
                "section",
                &["section"],
                &[K::Body],
                &[],
                &["/body/section[N]"],
                &[],
            ),
            // Story paragraphs are not listed children: unscoped `p`
            // queries stay in the main story (record 0319).
            identities(
                element(
                    K::Header,
                    "header",
                    &["header"],
                    &[K::Document],
                    &[],
                    &["/header[N]", "/header[@type=T]"],
                    super::parts::STORY_PROPS,
                ),
                &["type"],
            ),
            identities(
                element(
                    K::Footer,
                    "footer",
                    &["footer"],
                    &[K::Document],
                    &[],
                    &["/footer[N]", "/footer[@type=T]"],
                    super::parts::STORY_PROPS,
                ),
                &["type"],
            ),
            element(
                K::Footnotes,
                "footnotes",
                &["footnotes"],
                &[K::Document],
                &[K::Footnote],
                &["/footnotes"],
                &[],
            ),
            identities(
                element(
                    K::Footnote,
                    "footnote",
                    &["footnote"],
                    &[K::Document, K::Footnotes],
                    &[],
                    &[
                        "/footnotes/footnote[N]",
                        "/footnote[N]",
                        "/footnote[@footnoteId=N]",
                    ],
                    super::parts::NOTE_PROPS,
                ),
                &["footnoteId"],
            ),
            element(
                K::Endnotes,
                "endnotes",
                &["endnotes"],
                &[K::Document],
                &[K::Endnote],
                &["/endnotes"],
                &[],
            ),
            identities(
                element(
                    K::Endnote,
                    "endnote",
                    &["endnote"],
                    &[K::Document, K::Endnotes],
                    &[],
                    &[
                        "/endnotes/endnote[N]",
                        "/endnote[N]",
                        "/endnote[@endnoteId=N]",
                    ],
                    super::parts::NOTE_PROPS,
                ),
                &["endnoteId"],
            ),
            element(
                K::Comments,
                "comments",
                &["comments"],
                &[K::Document],
                &[K::Comment],
                &["/comments"],
                &[],
            ),
            identities(
                element(
                    K::Comment,
                    "comment",
                    &["comment"],
                    &[K::Comments],
                    &[],
                    &["/comments/comment[N]", "/comments/comment[@commentId=X]"],
                    COMMENT_PROPS,
                ),
                &["commentId"],
            ),
            element(
                K::Styles,
                "styles",
                &["styles"],
                &[K::Document],
                &[K::Style],
                &["/styles"],
                super::parts::STYLES_PROPS,
            ),
            element(
                K::Style,
                "style",
                &["style"],
                &[K::Styles],
                &[],
                &["/styles/<id or name>"],
                super::parts::STYLE_PROPS,
            ),
            element(
                K::Numbering,
                "numbering",
                &["numbering"],
                &[K::Document],
                &[K::AbstractNum, K::Num],
                &["/numbering"],
                super::parts::NUMBERING_PROPS,
            ),
            identities(
                element(
                    K::AbstractNum,
                    "abstractNum",
                    &["abstractnum"],
                    &[K::Numbering],
                    &[K::Level],
                    &["/numbering/abstractNum[N]", "/numbering/abstractNum[@id=N]"],
                    super::parts::ABSTRACT_NUM_PROPS,
                ),
                &["id"],
            ),
            identities(
                element(
                    K::Num,
                    "num",
                    &["num"],
                    &[K::Numbering],
                    &[],
                    &["/numbering/num[N]", "/numbering/num[@id=N]"],
                    super::parts::NUM_PROPS,
                ),
                &["id"],
            ),
            identities(
                element(
                    K::Level,
                    "level",
                    &["level", "lvl"],
                    &[K::AbstractNum],
                    &[],
                    &["/numbering/abstractNum[N]/level[L]"],
                    super::parts::LEVEL_PROPS,
                ),
                &["ilvl"],
            ),
            gap(
                K::Settings,
                "settings",
                &["settings"],
                &[K::Document],
                &["/settings"],
                &SETTINGS_GAP,
            ),
            gap(
                K::Theme,
                "theme",
                &["theme"],
                &[K::Document],
                &["/theme"],
                &THEME_GAP,
            ),
            identities(
                element(
                    K::Bookmark,
                    "bookmark",
                    &["bookmark"],
                    &[K::Document],
                    &[],
                    &["/bookmark[N]", "/bookmark[@name=X]"],
                    super::marks::BOOKMARK_PROPS,
                ),
                &["name"],
            ),
            element(
                K::Field,
                "field",
                &["field"],
                &[K::Document],
                &[],
                &["/field[N]"],
                super::marks::FIELD_PROPS,
            ),
            gap(
                K::FormField,
                "formfield",
                &["formfield"],
                &[K::Paragraph],
                &["/formfield[NAME]"],
                &FORM_FIELD_GAP,
            ),
            element(
                K::Picture,
                "picture",
                &["image", "img", "picture"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/picture[M]"],
                PICTURE_PROPS,
            ),
            element(
                K::Chart,
                "chart",
                &["chart"],
                &[K::Paragraph, K::Document],
                &[K::Series, K::Axis],
                &["/body/p[N]/chart[M]", "/chart[N]", "/body/chart[N]"],
                super::chart::CHART_PROPS,
            ),
            gap(
                K::Ole,
                "ole",
                &["embed", "object", "ole", "oleobject"],
                &[K::Paragraph],
                &["/body/p[N]/ole[M]"],
                &OLE_GAP,
            ),
            element(
                K::Toc,
                "toc",
                &["tableofcontents", "toc"],
                &[K::Document],
                &[],
                &["/toc[N]"],
                super::marks::TOC_PROPS,
            ),
            gap(
                K::Revision,
                "revision",
                &["revision"],
                &[K::Document],
                &["/revision[N]"],
                &REVISION_GAP,
            ),
            element(
                K::Watermark,
                "watermark",
                &["watermark"],
                &[K::Document],
                &[],
                &["/watermark"],
                super::marks::WATERMARK_PROPS,
            ),
            element(
                K::TextBox,
                "textbox",
                &["textbox"],
                &[K::Body],
                &[],
                &["/body/textbox[N]"],
                super::marks::TEXT_BOX_PROPS,
            ),
            gap(
                K::Shape,
                "shape",
                &["shape"],
                &[K::Paragraph],
                &["/body/p[N]/shape[M]"],
                &SHAPE_GAP,
            ),
            element(
                K::Break,
                "break",
                &["br", "break", "pagebreak"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/break[M]"],
                super::marks::BREAK_PROPS,
            ),
            element(
                K::Tab,
                "tab",
                &["tab"],
                &[K::Paragraph],
                &[],
                &["/body/p[N]/tab[M]"],
                super::marks::TAB_PROPS,
            ),
            element(
                K::SmartArt,
                "smartart",
                &["smartart"],
                &[K::Document],
                &[K::Node],
                &["/smartart[N]", "/body/smartart[N]"],
                super::smartart::SMARTART_PROPS,
            ),
            element(
                K::Series,
                "series",
                &["series"],
                &[K::Chart],
                &[],
                &["/body/p[N]/chart[M]/series[K]", "/chart[N]/series[K]"],
                super::chart::SERIES_PROPS,
            ),
            // `role` selects an axis (`axis[@role=value]`) and is also an
            // ordinary property, as on PPTX.
            identities(
                element(
                    K::Axis,
                    "axis",
                    &["axis"],
                    &[K::Chart],
                    &[],
                    &[
                        "/body/p[N]/chart[M]/axis[@role=ROLE]",
                        "/body/p[N]/chart[M]/axis[N]",
                        "/chart[N]/axis[@role=ROLE]",
                    ],
                    super::chart::AXIS_PROPS,
                ),
                &["role"],
            ),
            element(
                K::Node,
                "node",
                &["node"],
                &[K::SmartArt, K::Node],
                &[K::Node],
                &["/smartart[N]/node[K]", "/smartart[N]/node[K]/node[J]"],
                super::smartart::NODE_PROPS,
            ),
            // Not a listed table child: table views and child counts stay
            // the rows (record 0345).
            element(
                K::Column,
                "column",
                &["col", "column"],
                &[K::Table],
                &[],
                &["/body/tbl[N]/col[M]"],
                COLUMN_PROPS,
            ),
        ],
        gaps: &[
            // OfficeCLI elements this adapter does not list as kinds
            // (record 0339): each is a typed gap with its disposition.
            unlisted_gap(
                "bookmarkEnd",
                "0339-F2 (deferred-P1)",
                "a bookmark is one node over its start and end markers (record 0319); a standalone end marker is not addressed or added on its own",
            ),
            unlisted_gap(
                "permStart",
                "0339-F2 (deferred-P1)",
                "editing-permission ranges (w:permStart, w:permEnd) are kept but not projected as nodes",
            ),
            unlisted_gap(
                "ptab",
                "0339-F2 (deferred-P1)",
                "positional tabs (w:ptab) are kept in their runs but not projected as nodes",
            ),
            unlisted_gap(
                "fieldChar",
                "none: out of scope (record 0332)",
                "a complex field is one field node (begin, instruction, result and end); its field-character and instruction runs are not separate elements",
            ),
            unlisted_gap(
                "instrText",
                "none: out of scope (record 0332)",
                "a complex field is one field node (begin, instruction, result and end); its field-character and instruction runs are not separate elements",
            ),
            GapSpec {
                subject: Subject::Element("diagram"),
                code: Code::UnsupportedPath,
                follow_up: "none: out of scope (record 0332, design §10)",
                reason: "OfficeCLI's diagram renders Mermaid text into shapes or a picture; rendering is outside the unified API, and SmartArt is the smartart element",
            },
            GapSpec {
                subject: Subject::Property {
                    element: "sdt",
                    key: "type",
                },
                code: Code::UnsupportedCapability,
                follow_up: "R4x",
                reason: "the leaf reads a content control's id, tag and alias only",
            },
            GapSpec {
                subject: Subject::Property {
                    element: "sdt",
                    key: "lock",
                },
                code: Code::UnsupportedCapability,
                follow_up: "R4x",
                reason: "the leaf reads a content control's id, tag and alias only",
            },
            GapSpec {
                subject: Subject::Property {
                    element: "section",
                    key: "pageWidth",
                },
                code: Code::UnsupportedCapability,
                follow_up: "R4x",
                reason: "section properties are read through the section inventory, which the node view does not wire yet",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "set",
                    element: "paragraph",
                    context: "wrapped-or-unproven-selection",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W3",
                reason: "paragraphs in cells and content controls are edited through proven element paths (record 0251a); paragraphs in w:customXml wrappers, wrapped rows or cells, or markup-compatibility content selected outside runs have no proven raw position; alternate content inside runs keeps every position (record 0307)",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "set",
                    element: "run",
                    context: "wrapped-or-unproven-selection",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2",
                reason: "direct runs and direct runs of hyperlinks are edited (record 0251a); runs in tracked changes, smart tags, custom XML or other wrappers, and paragraphs with comment-reference runs, have no proven raw ordinal",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "set",
                    element: "paragraph/run",
                    context: "mixed-paragraph-run-fanout",
                },
                code: Code::UnsupportedCapability,
                follow_up: "M2 (homogeneous mutation fanout)",
                reason: "one DOCX mutation admits one target kind, so a paragraph/run union is refused before staging rather than applying a descendant ordinal to the wrong leaf",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "set",
                    element: "paragraph/run",
                    context: "managed-transaction",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W5",
                reason: "the formatting transaction has not proven edit admission and publication for execution-context-managed DOCX owners; the read projection remains available",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "set",
                    element: "paragraph/run",
                    context: "complex-leaf-content",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2 / record 0011 / record 0307 (text-box stories)",
                reason: "fields and other nested or unmodeled run content lack a proven raw ownership and address mapping for leaf edits; markup-compatibility alternate content in a run is one opaque object, so a text set over one whose branches hold text (a text box) is refused",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "add",
                    element: "paragraph/run",
                    context: "managed-transaction",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W5",
                reason: "structural insertion uses the unmanaged DOCX transaction path; execution-context-managed owners refuse the leaf splice until managed admission is proven",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "add",
                    element: "paragraph",
                    context: "wrapped-or-unproven-parent",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W3",
                reason: "paragraphs are added under the body, cells and block content controls (record 0251a); parents whose raw position is not proven are refused",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "add",
                    element: "run",
                    context: "wrapped-or-unproven-parent",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2",
                reason: "runs are added beside the direct runs of any proven paragraph (record 0251a); insertion inside a hyperlink and under wrapped paragraphs is refused",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "remove",
                    element: "paragraph/run",
                    context: "managed-transaction",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W5",
                reason: "structural removal uses the unmanaged DOCX transaction path; execution-context-managed owners refuse the leaf splice until managed admission is proven",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "remove",
                    element: "paragraph/run",
                    context: "wrapped-or-unproven-selection",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2 / L-DOCX-W3",
                reason: "paragraphs, runs, hyperlinks, tables, rows, cells and content controls are removed through proven element paths (record 0251a); wrapped selections, the last paragraph of a cell, the last cell of a row and the last row of a table are refused",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "remove",
                    element: "paragraph/run",
                    context: "complex-leaf-content",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2 / record 0011",
                reason: "paragraph and run dependency scans refuse open ranges, field interiors, tracked changes and story references rather than dropping their relationships",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "move",
                    element: "paragraph/run",
                    context: "managed-transaction",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W5",
                reason: "paragraph and direct-run movement use the unmanaged DOCX transaction path; execution-context-managed owners refuse the leaf splice until managed admission is proven",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "move",
                    element: "paragraph/run",
                    context: "wrapped-or-unproven-selection",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W3",
                reason: "body paragraphs, direct runs, tables, rows and nested paragraphs move through proven anchors (record 0251a); nested runs and wrapped selections have no proven anchor",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "move",
                    element: "paragraph/run",
                    context: "complex-leaf-content",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2 / record 0011",
                reason: "paragraph and direct-run movement preserves exact source fragments but refuses open ranges, field interiors, namespace changes, tracked changes and unmodeled dependencies",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "swap",
                    element: "paragraph/run",
                    context: "managed-transaction",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W5",
                reason: "paragraph and direct-run swapping compose unmanaged fragment moves; execution-context-managed owners refuse the leaf splice until managed admission is proven",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "swap",
                    element: "paragraph/run",
                    context: "nested-or-unproven-selection",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W3",
                reason: "only direct main-body paragraphs or direct runs can be swapped; nested and wrapped selections have no proven anchors",
            },
            GapSpec {
                subject: Subject::Operation {
                    verb: "swap",
                    element: "paragraph/run",
                    context: "complex-leaf-content",
                },
                code: Code::UnsupportedCapability,
                follow_up: "L-DOCX-W2 / record 0011",
                reason: "paragraph and direct-run swapping preserves exact source fragments but refuses field interiors, namespace changes, tracked changes and unmodeled dependencies",
            },
            GapSpec {
                subject: Subject::View("html"),
                code: Code::UnsupportedCapability,
                follow_up: "design §10",
                reason: "rendering is outside the unified layer",
            },
        ],
    }
};

check_format_spec!(SPEC);

const _: () = assert!(
    SPEC.elements.len() <= 128,
    "a query kind set holds 128 kinds"
);
