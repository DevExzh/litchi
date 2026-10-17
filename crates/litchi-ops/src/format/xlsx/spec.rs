//! The SpreadsheetML specialization of the shared workbook vocabulary.

use crate::format::workbook::spec as workbook;
use crate::node::spec::{
    Access as ReadAccess, ElementSpec, FormatSpec, GapSpec, PropSpec, Subject, check_format_spec,
};
use crate::registry::{Access, Alias, Constraint, Entry, Table};
use crate::value::{ColorContext, Length, LengthSpec, Tokens, ValueType};

pub(crate) use workbook::{
    CELL_PROPS, COLUMN_PROPS, CellKey, ColumnKey, Kind, NAMED_RANGE_PROPS, NamedRangeKey,
    ROW_PROPS, RowKey, SHEET_PROPS, SheetKey, TABLE_PROPS, TableKey, WORKBOOK_PROPS, WorkbookKey,
};

/// The one readable property on an XLSX range.
///
/// The shared workbook table leaves this property as a gap because the XLSB
/// adapter does not yet project range merge state.  XLSX resolves it from its
/// native merge index, so it has a format-local read table and key enum.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RangeKey {
    Merge,
}

impl RangeKey {
    const ALL: &'static [Self] = &[Self::Merge];

    pub(crate) fn at(index: usize) -> Option<Self> {
        Self::ALL.get(index).copied()
    }
}

pub(crate) const RANGE_PROPS: &[PropSpec] = &[PropSpec::new("merge", ValueType::Bool)];

/// Merge assignment admission is shared by cell, range and sheet targets.
/// `Scalar` admits typed booleans for range true/false and text for range
/// literals, `sweep`, and cell/sheet comma-separated lists.  The merge route
/// rejects numeric scalar variants after registry coercion and applies the
/// target-specific grammar.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MergeSetter {
    Merge,
}

pub(crate) const MERGE_WRITE_TABLE: Table<MergeSetter, 1, 0> = Table::new(
    [Entry {
        spec: PropSpec::new("merge", ValueType::Scalar),
        access: Access::Set(MergeSetter::Merge),
        constraint: Constraint::None,
        clear: true,
    }],
    [],
);

/// Read/capability metadata derived from the merge setter registry.
pub(crate) const MERGE_WRITE_PROPS: &[PropSpec] = &MERGE_WRITE_TABLE.readable();

/// XLSX cell underline values.  `none` remains distinct from a JSON null:
/// the former authors an explicit SpreadsheetML `<u val="none"/>`, while the
/// latter removes the local underline element.  The textual boolean and short token
/// spellings are the OfficeCLI cell aliases and are canonicalized here before
/// native font admission.
pub(super) static XLSX_UNDERLINE: Tokens = Tokens::new(
    "underline",
    &[
        "single",
        "double",
        "singleAccounting",
        "doubleAccounting",
        "none",
    ],
    &[
        ("0", 4),
        ("1", 0),
        ("dbl", 1),
        ("doubleacct", 3),
        ("false", 4),
        ("no", 4),
        ("off", 4),
        ("on", 0),
        ("singleacct", 2),
        ("true", 0),
        ("yes", 0),
    ],
);

/// Border line tokens accepted by the XLSX writer. The shared read table
/// omits `none` because an absent side is represented by a missing value;
/// authoring needs the token to clear one side while retaining the rest of a
/// border record. Token matching remains case-insensitive, as in the shared
/// vocabulary policy.
static XLSX_BORDER_LINE: Tokens = Tokens::new(
    "border style",
    &[
        "thin",
        "medium",
        "dashed",
        "dotted",
        "thick",
        "double",
        "hair",
        "mediumDashed",
        "dashDot",
        "mediumDashDot",
        "dashDotDot",
        "mediumDashDotDot",
        "slantDashDot",
        "none",
    ],
    &[],
);

/// XLSX workbook-tab assignments supported by the immutable write adapter.
///
/// `name` and `hidden` intentionally reuse the shared sheet property spelling
/// and value types.  The read projection remains [`SHEET_PROPS`]; this table
/// is the one admission source used by both the set adapter and capability
/// reporting.  Native tab editing remains authoritative for name validation,
/// duplicate-name checks and the active/last-visible-tab policy.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetSetter {
    Name,
    Hidden,
}

pub(crate) const SHEET_WRITE_TABLE: Table<SheetSetter, 2, 0> = Table::new(
    [
        Entry {
            spec: PropSpec::new("name", ValueType::Text),
            access: Access::SetAndAdd(SheetSetter::Name),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("hidden", ValueType::Bool),
            access: Access::SetAndAdd(SheetSetter::Hidden),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [],
);

/// Read-schema projection of the sheet write registry.
pub(crate) const SHEET_WRITE_PROPS: &[PropSpec] = &SHEET_WRITE_TABLE.readable();

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access = ReadAccess {
        get: false,
        set: false,
        add: false,
    };
    spec
}

/// Cell properties supported by the immutable XLSX write adapter.
///
/// The read projection deliberately continues to use the shared
/// [`CELL_PROPS`] table above. These entries are a write-only registry seam:
/// `value` accepts the registry scalar union (text, bool, an integer within
/// the signed 64-bit range, or a finite number), `formula` accepts an explicit
/// formula expression without `=`, and
/// `clear` accepts only `true`. A JSON `null` on `value` or `formula` is also
/// an explicit clear, as allowed by the common property contract. `font.bold`
/// and `font.italic` accept boolean assignments or null to remove the local
/// font flag. `strike` accepts the OfficeCLI `strikethrough` and
/// `font.strike` and `font.strikethrough` aliases. `underline` accepts the
/// five SpreadsheetML styles, the `font.underline` alias and OfficeCLI
/// boolean/short-token aliases; null removes the authored underline element.
/// `font.name` (aliases `font` and `fontname`) accepts non-empty XML-safe text;
/// `font.size` (aliases `size` and `fontsize`) accepts 0.5 through 409 points;
/// and `superscript`/`subscript` (with their `font.*` aliases) accept boolean
/// values or null. `font.color` (and the `color` alias) accepts the checked
/// SpreadsheetML RGB/theme vocabulary or null. The remaining native font
/// children use their explicit `font.*` names: `outline`, `shadow`,
/// `condense`, and `extend` are booleans; `charset` is an unsigned byte
/// (`0..=255`); `family` is the SpreadsheetML font-family class
/// (`0..=5`); and `scheme` accepts `none`, `major`, or `minor`. These
/// long-tail names intentionally have no bare aliases because OfficeCLI
/// preserves them as typed XML children rather than advertising shorthand
/// setter keys. Every field accepts null to remove its authored child.
/// Alignment uses the shared SpreadsheetML vocabularies. The canonical keys
/// are `alignment.horizontal`, `alignment.vertical`, `alignment.wrapText`,
/// `alignment.textRotation`, `alignment.indent`, `alignment.shrinkToFit`, and
/// `alignment.readingOrder`; OfficeCLI's top-level aliases (`align`, `halign`,
/// `valign`, `wrap`, `rotation`, `textrotation`, `indent`, `shrink`,
/// `readingorder`, `direction`, and `dir`) plus the dotted `alignment.rotation`
/// and `alignment.shrink` spellings resolve to those entries. Rotation accepts
/// `0..=180` and the
/// stacked sentinel `255`; the wire contextual sentinel `254` is refused to
/// match OfficeCLI's authoring policy. Indent accepts `0..=255`, and reading
/// order accepts `context`, `ltr`, or `rtl`. Alignment assignments can
/// accompany font, protection, quote-prefix and payload assignments in one
/// native edit. `alignment.relativeIndent` uses the signed SpreadsheetML
/// `Int32Value` domain in the local Open XML schema; it is therefore admitted
/// across the complete `i32` range rather than narrowed to the usual
/// non-negative `indent` range. `alignment.justifyLastLine` is a boolean.
/// `protection.locked`, `protection.hidden`, and `quotePrefix` accept boolean
/// values or null to remove their authored attributes. The OfficeCLI aliases
/// `locked` and `formulahidden` are the only shorthand aliases for the two
/// protection entries.
/// `numberformat` accepts a validated Excel format code or one of OfficeCLI's
/// whole-value friendly aliases (`general`, `number`, `comma`, `currency`,
/// `accounting`, `percent`, `percentage`, `scientific`, `text`, `date`, `time`,
/// and `datetime`). `numfmt` and `format` are its only setter aliases; custom
/// format code spelling is retained, while aliases are normalized before the
/// native style editor sees them.
/// `fill`/`fillBg` accept checked SpreadsheetML RGB or theme colors. The
/// write-only `fill` entry uses text admission because it also accepts
/// OfficeCLI's bounded multi-stop gradient spellings
/// (`C1-C2[...][-degree]`, `radial:C1-C2[...]`, and
/// `gradient;C1;C2[;...];degree`). `fill` and `fillBg` preserve an existing
/// non-solid pattern when the caller omits `fillPattern`, while `fillPattern`
/// replaces the pattern and clears omitted colors. A gradient is a complete
/// fill replacement and cannot be combined with `fillPattern` or `fillBg` in
/// one request. `bgcolor` and `bg` are the OfficeCLI aliases for `fill`.
/// Border setters use the shared read names. `border` and `border.all` fan
/// out a line style to the four physical sides, while `border.color` and
/// `border.all.color` fan out a color. Side colors preserve an existing line;
/// a color-only request on a side without a line is refused rather than
/// authoring an invisible color. `none` clears a side's line while preserving
/// an authored color unless that color is also explicitly cleared with null;
/// null clears an authored color or direction flag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellSetter {
    Value,
    Formula,
    Clear,
    Bold,
    Italic,
    Strike,
    Underline,
    FontName,
    FontSize,
    Superscript,
    Subscript,
    FontColor,
    FontOutline,
    FontShadow,
    FontCondense,
    FontExtend,
    FontCharset,
    FontFamily,
    FontScheme,
    AlignmentHorizontal,
    AlignmentVertical,
    AlignmentWrapText,
    AlignmentTextRotation,
    AlignmentIndent,
    AlignmentShrinkToFit,
    AlignmentReadingOrder,
    AlignmentRelativeIndent,
    AlignmentJustifyLastLine,
    ProtectionLocked,
    ProtectionHidden,
    QuotePrefix,
    NumberFormat,
    Fill,
    FillPattern,
    FillBg,
    BorderAll,
    BorderAllColor,
    BorderLeft,
    BorderLeftColor,
    BorderRight,
    BorderRightColor,
    BorderTop,
    BorderTopColor,
    BorderBottom,
    BorderBottomColor,
    BorderDiagonal,
    BorderDiagonalColor,
    BorderDiagonalUp,
    BorderDiagonalDown,
    /// The cell type (record 0338): `string`, `number`, `boolean`, `date`,
    /// `error` or `richtext`.
    Type,
    /// Rich-text runs, a JSON array (record 0338).
    Runs,
    /// The address of an added cell (record 0338); add-only.
    Ref,
}

/// The cell `type` an `add` or `set` writes (record 0338), with OfficeCLI's
/// short spellings (`ExcelHandler.Set.Cells.cs`).
pub(crate) static CELL_TYPE: Tokens = Tokens::new(
    "cell type",
    &["string", "number", "boolean", "date", "error", "richtext"],
    &[("bool", 2), ("err", 4), ("num", 1), ("rich", 5), ("str", 0)],
);

pub(crate) const CELL_WRITE_TABLE: Table<CellSetter, 54, 42> = Table::new(
    [
        Entry {
            spec: write_only("value", ValueType::Scalar),
            access: Access::SetAndAdd(CellSetter::Value),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("formula", ValueType::Text),
            access: Access::SetAndAdd(CellSetter::Formula),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("clear", ValueType::Bool),
            access: Access::Set(CellSetter::Clear),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: PropSpec::new("font.bold", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::Bold),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.italic", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::Italic),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("strike", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::Strike),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("underline", ValueType::Token(&XLSX_UNDERLINE)),
            access: Access::SetAndAdd(CellSetter::Underline),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.name", ValueType::Text),
            access: Access::SetAndAdd(CellSetter::FontName),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.size", ValueType::Length(LengthSpec::FONT_SIZE)),
            access: Access::SetAndAdd(CellSetter::FontSize),
            constraint: Constraint::Length {
                min: Length::ZERO,
                max: Length::from_points(409),
            },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("superscript", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::Superscript),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("subscript", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::Subscript),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.color", ValueType::Color(ColorContext::Spreadsheet)),
            access: Access::SetAndAdd(CellSetter::FontColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.outline", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::FontOutline),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.shadow", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::FontShadow),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.condense", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::FontCondense),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.extend", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::FontExtend),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.charset", ValueType::Int),
            access: Access::SetAndAdd(CellSetter::FontCharset),
            constraint: Constraint::Int { min: 0, max: 255 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.family", ValueType::Int),
            access: Access::SetAndAdd(CellSetter::FontFamily),
            constraint: Constraint::Int { min: 0, max: 5 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("font.scheme", ValueType::Token(&workbook::FONT_SCHEME)),
            access: Access::SetAndAdd(CellSetter::FontScheme),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "alignment.horizontal",
                ValueType::Token(&workbook::HORIZONTAL),
            ),
            access: Access::SetAndAdd(CellSetter::AlignmentHorizontal),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.vertical", ValueType::Token(&workbook::VERTICAL)),
            access: Access::SetAndAdd(CellSetter::AlignmentVertical),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.wrapText", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::AlignmentWrapText),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.textRotation", ValueType::Int),
            access: Access::SetAndAdd(CellSetter::AlignmentTextRotation),
            // OfficeCLI accepts 0..=180 and the stacked sentinel 255.  The
            // discontiguous 181..=254 gap is rejected during admission.
            constraint: Constraint::Int { min: 0, max: 255 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.indent", ValueType::Int),
            access: Access::SetAndAdd(CellSetter::AlignmentIndent),
            constraint: Constraint::Int { min: 0, max: 255 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.shrinkToFit", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::AlignmentShrinkToFit),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "alignment.readingOrder",
                ValueType::Token(&workbook::READING_ORDER),
            ),
            access: Access::SetAndAdd(CellSetter::AlignmentReadingOrder),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.relativeIndent", ValueType::Int),
            access: Access::SetAndAdd(CellSetter::AlignmentRelativeIndent),
            // The Open XML SDK schema declares this as Int32Value and adds no
            // narrower validator.  The adapter performs the i64 -> i32 check
            // after registry admission.
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("alignment.justifyLastLine", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::AlignmentJustifyLastLine),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("protection.locked", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::ProtectionLocked),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("protection.hidden", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::ProtectionHidden),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("quotePrefix", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::QuotePrefix),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("numberformat", ValueType::Text),
            access: Access::SetAndAdd(CellSetter::NumberFormat),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("fill", ValueType::Text),
            access: Access::SetAndAdd(CellSetter::Fill),
            constraint: Constraint::NotEmpty,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("fillPattern", ValueType::Token(&workbook::FILL_PATTERN)),
            access: Access::SetAndAdd(CellSetter::FillPattern),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("fillBg", ValueType::Color(ColorContext::Spreadsheet)),
            access: Access::SetAndAdd(CellSetter::FillBg),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderAll),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.all", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderAll),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.color", ValueType::Color(ColorContext::Spreadsheet)),
            access: Access::SetAndAdd(CellSetter::BorderAllColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.all.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderAllColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.left", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderLeft),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.left.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderLeftColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.right", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderRight),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.right.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderRightColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.top", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderTop),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.top.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderTopColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.bottom", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderBottom),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.bottom.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderBottomColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.diagonal", ValueType::Token(&XLSX_BORDER_LINE)),
            access: Access::SetAndAdd(CellSetter::BorderDiagonal),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new(
                "border.diagonal.color",
                ValueType::Color(ColorContext::Spreadsheet),
            ),
            access: Access::SetAndAdd(CellSetter::BorderDiagonalColor),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.diagonalUp", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::BorderDiagonalUp),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("border.diagonalDown", ValueType::Bool),
            access: Access::SetAndAdd(CellSetter::BorderDiagonalDown),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: PropSpec::new("type", ValueType::Token(&CELL_TYPE)),
            access: Access::SetAndAdd(CellSetter::Type),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: write_only("runs", ValueType::Text),
            access: Access::SetAndAdd(CellSetter::Runs),
            constraint: Constraint::None,
            clear: false,
        },
        Entry {
            spec: write_only("ref", ValueType::Text),
            access: Access::Add(CellSetter::Ref),
            constraint: Constraint::None,
            clear: false,
        },
    ],
    [
        // OfficeCLI's cell handler takes `text` as another name of `value`
        // (`ExcelHandler.Set.Cells.cs`, `case "value" or "text"`), although
        // its help schema lists only `value` (record 0269).
        Alias {
            name: "text",
            entry: 0,
        },
        // OfficeCLI's documented aliases (`cell.json`).
        Alias {
            name: "bold",
            entry: 3,
        },
        Alias {
            name: "italic",
            entry: 4,
        },
        Alias {
            name: "font.strike",
            entry: 5,
        },
        Alias {
            name: "font.underline",
            entry: 6,
        },
        Alias {
            name: "font.strikethrough",
            entry: 5,
        },
        Alias {
            name: "strikethrough",
            entry: 5,
        },
        Alias {
            name: "font",
            entry: 7,
        },
        Alias {
            name: "fontname",
            entry: 7,
        },
        Alias {
            name: "size",
            entry: 8,
        },
        Alias {
            name: "fontsize",
            entry: 8,
        },
        Alias {
            name: "font.superscript",
            entry: 9,
        },
        Alias {
            name: "font.subscript",
            entry: 10,
        },
        Alias {
            name: "color",
            entry: 11,
        },
        Alias {
            name: "halign",
            entry: 19,
        },
        Alias {
            name: "align",
            entry: 19,
        },
        Alias {
            name: "valign",
            entry: 20,
        },
        Alias {
            name: "wrap",
            entry: 21,
        },
        Alias {
            name: "wraptext",
            entry: 21,
        },
        Alias {
            name: "rotation",
            entry: 22,
        },
        Alias {
            name: "textrotation",
            entry: 22,
        },
        Alias {
            name: "alignment.rotation",
            entry: 22,
        },
        Alias {
            name: "indent",
            entry: 23,
        },
        Alias {
            name: "shrinktofit",
            entry: 24,
        },
        Alias {
            name: "shrink",
            entry: 24,
        },
        Alias {
            name: "alignment.shrink",
            entry: 24,
        },
        Alias {
            name: "readingorder",
            entry: 25,
        },
        Alias {
            name: "direction",
            entry: 25,
        },
        Alias {
            name: "dir",
            entry: 25,
        },
        Alias {
            name: "locked",
            entry: 28,
        },
        Alias {
            name: "formulahidden",
            entry: 29,
        },
        Alias {
            name: "numfmt",
            entry: 31,
        },
        Alias {
            name: "format",
            entry: 31,
        },
        Alias {
            name: "bgcolor",
            entry: 32,
        },
        Alias {
            name: "bg",
            entry: 32,
        },
        Alias {
            name: "border.all.style",
            entry: 36,
        },
        Alias {
            name: "border.left.style",
            entry: 39,
        },
        Alias {
            name: "border.right.style",
            entry: 41,
        },
        Alias {
            name: "border.top.style",
            entry: 43,
        },
        Alias {
            name: "border.bottom.style",
            entry: 45,
        },
        Alias {
            name: "border.diagonal.style",
            entry: 47,
        },
        // OfficeCLI's alias of `ref` (`cell.json`).
        Alias {
            name: "address",
            entry: 53,
        },
    ],
);

/// Read-schema projection of the write registry for capability reporting.
pub(crate) const CELL_WRITE_PROPS: &[PropSpec] = &CELL_WRITE_TABLE.readable();

/// Cell capability projection: all shared read properties in their existing
/// order, with matching registry properties upgraded to writable and every
/// other registry entry (`value`, `clear`, the border shorthands) appended
/// write-only.  The XLSB adapter continues to consume only the
/// shared `CELL_PROPS` slice, so this write metadata cannot alter its read
/// semantics.
/// Whether `name` is a shared readable cell property.
const fn cell_reads(name: &str) -> bool {
    let mut index = 0;
    while index < CELL_PROPS.len() {
        if same_name(CELL_PROPS[index].name, name) {
            return true;
        }
        index += 1;
    }
    false
}

/// The cell write registry entries that are not readable cell properties
/// (`value`, `clear` and the border shorthands): appended write-only.
const fn cell_write_only_count() -> usize {
    let mut count = 0;
    let mut index = 0;
    while index < CELL_WRITE_PROPS.len() {
        if !cell_reads(CELL_WRITE_PROPS[index].name) {
            count += 1;
        }
        index += 1;
    }
    count
}

/// The cell hyperlink setters (record 0320): `link` (alias `url`),
/// `tooltip` (alias `screentip`) and `display`, written through the cell
/// edit; `null`, an empty link or `none` removes the cell's hyperlink.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellLinkSetter {
    Link,
    Tooltip,
    Display,
}

pub(crate) const CELL_LINK_TABLE: Table<CellLinkSetter, 3, 2> = Table::new(
    [
        Entry {
            spec: write_only("link", ValueType::Text),
            access: Access::Set(CellLinkSetter::Link),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("tooltip", ValueType::Text),
            access: Access::Set(CellLinkSetter::Tooltip),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("display", ValueType::Text),
            access: Access::Set(CellLinkSetter::Display),
            constraint: Constraint::None,
            clear: true,
        },
    ],
    [
        Alias {
            name: "screentip",
            entry: 1,
        },
        Alias {
            name: "url",
            entry: 0,
        },
    ],
);

const CELL_LINK_PROPS: &[PropSpec] = &CELL_LINK_TABLE.readable();

const CELL_CAPABILITY_LEN: usize =
    CELL_PROPS.len() + cell_write_only_count() + CELL_LINK_PROPS.len();

const fn cell_capability_array() -> [PropSpec; CELL_CAPABILITY_LEN] {
    let mut props = [PropSpec::new("", ValueType::Text); CELL_CAPABILITY_LEN];
    let mut index = 0;
    while index < CELL_PROPS.len() {
        let mut prop = CELL_PROPS[index];
        let mut write_index = 0;
        while write_index < CELL_WRITE_PROPS.len() {
            let write = CELL_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name) {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    // `merge` is already a readable cell property, so only upgrade its set
    // flag here.  Keep the read type (`string`) independent of the setter's
    // scalar admission type.
    let mut merge_index = 0;
    while merge_index < MERGE_WRITE_PROPS.len() {
        let write = MERGE_WRITE_PROPS[merge_index];
        let mut prop_index = 0;
        while prop_index < props.len() {
            if same_name(props[prop_index].name, write.name) {
                props[prop_index].access.set = write.access.set;
                props[prop_index].access.add = write.access.add;
                props[prop_index].admission = write.admission;
            }
            prop_index += 1;
        }
        merge_index += 1;
    }
    let mut next = CELL_PROPS.len();
    let mut write_index = 0;
    while write_index < CELL_WRITE_PROPS.len() {
        let mut write = CELL_WRITE_PROPS[write_index];
        if !cell_reads(write.name) {
            write.access.get = false;
            props[next] = write;
            next += 1;
        }
        write_index += 1;
    }
    let mut link_index = 0;
    while link_index < CELL_LINK_PROPS.len() {
        props[next] = CELL_LINK_PROPS[link_index];
        next += 1;
        link_index += 1;
    }
    props
}

const fn same_name(left: &str, right: &str) -> bool {
    let left = left.as_bytes();
    let right = right.as_bytes();
    if left.len() != right.len() {
        return false;
    }
    let mut index = 0;
    while index < left.len() {
        if left[index] != right[index] {
            return false;
        }
        index += 1;
    }
    true
}

pub(crate) static CELL_CAPABILITY_PROPS: &[PropSpec] = &cell_capability_array();

const fn range_capability_array() -> [PropSpec; RANGE_PROPS.len() + CELL_WRITE_PROPS.len()] {
    let mut props =
        [PropSpec::new("", ValueType::Text); RANGE_PROPS.len() + CELL_WRITE_PROPS.len()];
    let mut index = 0;
    while index < RANGE_PROPS.len() {
        let mut prop = RANGE_PROPS[index];
        if same_name(prop.name, MERGE_WRITE_PROPS[0].name) {
            prop.access.set = MERGE_WRITE_PROPS[0].access.set;
            prop.access.add = MERGE_WRITE_PROPS[0].access.add;
            prop.admission = MERGE_WRITE_PROPS[0].admission;
        }
        props[index] = prop;
        index += 1;
    }
    let mut write_index = 0;
    let write_props = CELL_WRITE_TABLE.readable();
    while write_index < write_props.len() {
        let mut prop = write_props[write_index];
        prop.access.get = false;
        // A range is not added (record 0338: cell properties are addable).
        prop.access.add = false;
        props[RANGE_PROPS.len() + write_index] = prop;
        write_index += 1;
    }
    props
}

/// Range assignments expose the native merge readback plus the cell setter
/// properties broadcast across a range.
pub(crate) static RANGE_CAPABILITY_PROPS: &[PropSpec] = &range_capability_array();

/// Sheet capability projection: the shared read table remains in its existing
/// order, while `name` and `hidden` gain their write flags from
/// [`SHEET_WRITE_TABLE`].  `visibility` stays read-only until the native
/// very-hidden policy is exposed by a later adapter.  Keeping this merge
/// derived from the write table prevents capability metadata and admission
/// from drifting apart.
/// The sheet-level auto-filter setter (record 0320): `autoFilter=A1:B5`
/// sets the worksheet's auto-filter range; `null`, `none` or `false`
/// removes it. The filter itself is the `autofilter` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum SheetFilterSetter {
    AutoFilter,
}

pub(crate) const SHEET_FILTER_TABLE: Table<SheetFilterSetter, 1, 0> = Table::new(
    [Entry {
        spec: write_only("autoFilter", ValueType::Text),
        access: Access::Set(SheetFilterSetter::AutoFilter),
        constraint: Constraint::None,
        clear: true,
    }],
    [],
);

const fn sheet_capability_array() -> [PropSpec; SHEET_PROPS.len() + 2] {
    let mut props = [PropSpec::new("", ValueType::Text); SHEET_PROPS.len() + 2];
    let mut index = 0;
    while index < SHEET_PROPS.len() {
        let mut prop = SHEET_PROPS[index];
        let mut write_index = 0;
        while write_index < SHEET_WRITE_PROPS.len() {
            let write = SHEET_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name) {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    let mut merge = MERGE_WRITE_PROPS[0];
    merge.access.get = false;
    props[SHEET_PROPS.len()] = merge;
    props[SHEET_PROPS.len() + 1] = SHEET_FILTER_TABLE.readable()[0];
    props
}

pub(crate) static SHEET_CAPABILITY_PROPS: &[PropSpec] = &sheet_capability_array();

/// The workbook root's capability table: its document-property keys are
/// settable on `/` through the metadata writer (record 0338).
pub(crate) static WORKBOOK_CAPABILITY_PROPS: [PropSpec; WORKBOOK_PROPS.len()] =
    crate::format::metadata::root_capability(WORKBOOK_PROPS);

/// Row assignments supported by the native worksheet layout editor.
///
/// The wire-facing read property remains `outlineLevel`, while OfficeCLI's
/// setter vocabulary uses `outline` and accepts `outlineLevel`/`group` as
/// aliases.  `outline` is therefore write-only in the capability projection;
/// the existing `outlineLevel` read property is deliberately left in place.
/// Native reset operations are represented by `null`: they remove the
/// explicit height/custom-height or outline marker, show the row, or expand
/// it, respectively.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RowSetter {
    Height,
    Hidden,
    Outline,
    Collapsed,
}

pub(crate) const ROW_WRITE_TABLE: Table<RowSetter, 4, 2> = Table::new(
    [
        Entry {
            spec: PropSpec::new("height", ROW_PROPS[0].ty),
            access: Access::Set(RowSetter::Height),
            constraint: Constraint::Length {
                min: Length::from_points(0),
                max: Length::from_points(409),
            },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("hidden", ROW_PROPS[1].ty),
            access: Access::Set(RowSetter::Hidden),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("outline", ROW_PROPS[2].ty),
            access: Access::Set(RowSetter::Outline),
            constraint: Constraint::Int { min: 0, max: 7 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("collapsed", ROW_PROPS[3].ty),
            access: Access::Set(RowSetter::Collapsed),
            constraint: Constraint::None,
            clear: true,
        },
    ],
    [
        Alias {
            name: "group",
            entry: 2,
        },
        Alias {
            name: "outlinelevel",
            entry: 2,
        },
    ],
);

/// Read-schema projection of the row write registry.
pub(crate) const ROW_WRITE_PROPS: &[PropSpec] = &ROW_WRITE_TABLE.readable();

/// Column assignments supported by the native worksheet layout editor.
///
/// Widths use the native SpreadsheetML character-unit number domain
/// (`0..=255`).  `null` resets width/custom-width, shows the column, removes
/// its outline marker, or expands it, as appropriate.  `outline` follows the
/// same setter spelling and aliases as the row registry while the read
/// projection continues to expose `outlineLevel`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ColumnSetter {
    Width,
    Hidden,
    Outline,
    Collapsed,
}

pub(crate) const COLUMN_WRITE_TABLE: Table<ColumnSetter, 4, 2> = Table::new(
    [
        Entry {
            spec: PropSpec::new("width", COLUMN_PROPS[0].ty),
            access: Access::Set(ColumnSetter::Width),
            constraint: Constraint::Number {
                min: 0.0,
                max: 255.0,
            },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("hidden", COLUMN_PROPS[1].ty),
            access: Access::Set(ColumnSetter::Hidden),
            constraint: Constraint::None,
            clear: true,
        },
        Entry {
            spec: write_only("outline", COLUMN_PROPS[3].ty),
            access: Access::Set(ColumnSetter::Outline),
            constraint: Constraint::Int { min: 0, max: 7 },
            clear: true,
        },
        Entry {
            spec: PropSpec::new("collapsed", COLUMN_PROPS[4].ty),
            access: Access::Set(ColumnSetter::Collapsed),
            constraint: Constraint::None,
            clear: true,
        },
    ],
    [
        Alias {
            name: "group",
            entry: 2,
        },
        Alias {
            name: "outlinelevel",
            entry: 2,
        },
    ],
);

/// Read-schema projection of the column write registry.
pub(crate) const COLUMN_WRITE_PROPS: &[PropSpec] = &COLUMN_WRITE_TABLE.readable();

const fn row_capability_array() -> [PropSpec; ROW_PROPS.len() + 1] {
    let mut props = [PropSpec::new("", ValueType::Text); ROW_PROPS.len() + 1];
    let mut index = 0;
    while index < ROW_PROPS.len() {
        let mut prop = ROW_PROPS[index];
        let mut write_index = 0;
        while write_index < ROW_WRITE_PROPS.len() {
            let write = ROW_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name)
                || (same_name(prop.name, "outlineLevel") && same_name(write.name, "outline"))
            {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props[ROW_PROPS.len()] = ROW_WRITE_PROPS[2];
    props
}

/// Row capability projection derived from the row write registry.
pub(crate) static ROW_CAPABILITY_PROPS: &[PropSpec] = &row_capability_array();

const fn column_capability_array() -> [PropSpec; COLUMN_PROPS.len() + 1] {
    let mut props = [PropSpec::new("", ValueType::Text); COLUMN_PROPS.len() + 1];
    let mut index = 0;
    while index < COLUMN_PROPS.len() {
        let mut prop = COLUMN_PROPS[index];
        let mut write_index = 0;
        while write_index < COLUMN_WRITE_PROPS.len() {
            let write = COLUMN_WRITE_PROPS[write_index];
            if same_name(prop.name, write.name)
                || (same_name(prop.name, "outlineLevel") && same_name(write.name, "outline"))
            {
                prop.access.set = write.access.set;
                prop.access.add = write.access.add;
                prop.admission = write.admission;
            }
            write_index += 1;
        }
        props[index] = prop;
        index += 1;
    }
    props[COLUMN_PROPS.len()] = COLUMN_WRITE_PROPS[2];
    props
}

/// Column capability projection derived from the column write registry.
pub(crate) static COLUMN_CAPABILITY_PROPS: &[PropSpec] = &column_capability_array();

const fn with_props(
    element: &'static ElementSpec<Kind>,
    props: &'static [PropSpec],
) -> ElementSpec<Kind> {
    ElementSpec {
        kind: element.kind,
        name: element.name,
        aliases: element.aliases,
        parents: element.parents,
        children: element.children,
        identities: element.identities,
        paths: element.paths,
        props,
        gap: element.gap,
    }
}

const fn is_range_merge_gap(gap: &GapSpec) -> bool {
    match gap.subject {
        Subject::Property { element, key } => {
            // Cell hyperlinks and the sheet auto-filter are write-only
            // setters (record 0320); their read gaps stay.
            same_name(element, "range") && same_name(key, "merge")
        },
        _ => false,
    }
}

const fn removed_gap_count() -> usize {
    let source = workbook::SPEC.gaps;
    let mut count = 0;
    let mut index = 0;
    while index < source.len() {
        if is_range_merge_gap(&source[index]) {
            count += 1;
        }
        index += 1;
    }
    count
}

/// The shared gaps XLSX keeps, plus its gradient-fill gap and the element
/// gaps of record 0339.
const XLSX_GAP_COUNT: usize =
    workbook::SPEC.gaps.len() - removed_gap_count() + 1 + XLSX_ELEMENT_GAPS.len();

/// OfficeCLI's heuristic table node and SmartArt (record 0339).
const XLSX_ELEMENT_GAPS: [GapSpec; 2] = [
    GapSpec {
        subject: Subject::Element("detectedtable"),
        code: crate::error::Code::UnsupportedCapability,
        follow_up: "0339-F5 (deferred-P1): header-row table detection in queries",
        reason: "OfficeCLI's detectedtable is a header-row block found heuristically in unstructured cells and reported by table queries; Litchi's queries return declared tables only",
    },
    GapSpec {
        subject: Subject::Element("smartart"),
        code: crate::error::Code::UnsupportedCapability,
        follow_up: "0339-F6 (deferred-P1): a worksheet-drawing diagram model in litchi-xlsx",
        reason: "litchi-xlsx reads and writes pictures, shapes and charts in worksheet drawings but has no diagram model; a SmartArt frame is kept as it is and is not a node",
    },
];

const fn xlsx_gaps() -> [GapSpec; XLSX_GAP_COUNT] {
    let source = workbook::SPEC.gaps;
    let mut result = [source[0]; XLSX_GAP_COUNT];
    let mut source_index = 0;
    let mut result_index = 0;
    while source_index < source.len() {
        let gap = source[source_index];
        if !is_range_merge_gap(&gap) {
            result[result_index] = gap;
            result_index += 1;
        }
        source_index += 1;
    }
    result[result_index] = GapSpec {
        subject: Subject::Operation {
            verb: "get",
            element: "cell",
            context: "gradient-fill-projection",
        },
        code: crate::error::Code::UnsupportedCapability,
        follow_up: "M1 XLSX lossless scalar gradient property projection",
        reason: "ordinary fill readback follows OfficeCLI RGB stop text and integer degrees; complete geometry, stop positions and native colors remain available through Node::cell_format",
    };
    let mut extra = 0;
    while extra < XLSX_ELEMENT_GAPS.len() {
        result[result_index + 1 + extra] = XLSX_ELEMENT_GAPS[extra];
        extra += 1;
    }
    result
}

static XLSX_GAPS: [GapSpec; XLSX_GAP_COUNT] = xlsx_gaps();

/// The chart elements the shared workbook vocabulary lists as gaps, read
/// and written by the XLSX adapter (record 0270).
const XLSX_CHART: ElementSpec<Kind> = ElementSpec {
    props: super::chart::CHART_PROPS,
    gap: None,
    ..workbook::SPEC.elements[9]
};
const XLSX_SERIES: ElementSpec<Kind> = ElementSpec {
    props: super::chart::SERIES_PROPS,
    gap: None,
    ..workbook::SPEC.elements[10]
};
const XLSX_AXIS: ElementSpec<Kind> = ElementSpec {
    props: super::chart::AXIS_PROPS,
    identities: &["role"],
    gap: None,
    paths: &[
        "/<sheet>/chart[N]/axis[@role=<role>]",
        "/<sheet>/chart[N]/axis[N]",
    ],
    ..workbook::SPEC.elements[11]
};

/// A shared element gap XLSX reads and writes as a worksheet feature
/// (record 0320).
const fn feature(index: usize, props: &'static [PropSpec]) -> ElementSpec<Kind> {
    ElementSpec {
        props,
        gap: None,
        ..workbook::SPEC.elements[index]
    }
}

// Keep the shared workbook topology and all non-XLSX-specific gaps, replacing
// only the range property table and its former capability gap.  XLSB retains
// the shared range gap in its own local SPEC.
static XLSX_ELEMENTS: [ElementSpec<Kind>; 26] = [
    with_props(
        &workbook::SPEC.elements[0],
        workbook::SPEC.elements[0].props,
    ),
    with_props(
        &workbook::SPEC.elements[1],
        workbook::SPEC.elements[1].props,
    ),
    with_props(
        &workbook::SPEC.elements[2],
        workbook::SPEC.elements[2].props,
    ),
    with_props(&workbook::SPEC.elements[3], RANGE_PROPS),
    with_props(
        &workbook::SPEC.elements[4],
        workbook::SPEC.elements[4].props,
    ),
    with_props(
        &workbook::SPEC.elements[5],
        workbook::SPEC.elements[5].props,
    ),
    with_props(
        &workbook::SPEC.elements[6],
        workbook::SPEC.elements[6].props,
    ),
    with_props(
        &workbook::SPEC.elements[7],
        workbook::SPEC.elements[7].props,
    ),
    feature(8, super::features::COMMENT_PROPS),
    XLSX_CHART,
    XLSX_SERIES,
    XLSX_AXIS,
    with_props(
        &workbook::SPEC.elements[12],
        workbook::SPEC.elements[12].props,
    ),
    with_props(
        &workbook::SPEC.elements[13],
        workbook::SPEC.elements[13].props,
    ),
    with_props(
        &workbook::SPEC.elements[14],
        workbook::SPEC.elements[14].props,
    ),
    feature(15, super::features::VALIDATION_PROPS),
    feature(16, super::features::CF_PROPS),
    feature(17, super::features::AUTOFILTER_PROPS),
    feature(18, super::features::HYPERLINK_PROPS),
    with_props(
        &workbook::SPEC.elements[19],
        workbook::SPEC.elements[19].props,
    ),
    feature(20, super::features::PICTURE_PROPS),
    feature(21, super::features::SHAPE_PROPS),
    with_props(
        &workbook::SPEC.elements[22],
        workbook::SPEC.elements[22].props,
    ),
    feature(23, super::features::ROW_BREAK_PROPS),
    feature(24, super::features::COL_BREAK_PROPS),
    // Rich-text runs of a cell (record 0338).
    ElementSpec {
        props: super::runs::RUN_PROPS,
        gap: None,
        paths: &["/<sheet>/<A1>/run[N]"],
        ..workbook::SPEC.elements[25]
    },
];

/// XLSX's format specification, with native range merge readback.  The
/// corresponding XLSB specification intentionally keeps the shared range
/// property gap until its binary projection is implemented.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: workbook::SPEC.root,
    elements: &XLSX_ELEMENTS,
    gaps: &XLSX_GAPS,
};

check_format_spec!(SPEC);

#[cfg(test)]
mod tests {
    use super::{
        CELL_CAPABILITY_PROPS, COLUMN_CAPABILITY_PROPS, COLUMN_WRITE_TABLE, MERGE_WRITE_TABLE,
        RANGE_CAPABILITY_PROPS, RANGE_PROPS, ROW_CAPABILITY_PROPS, ROW_WRITE_TABLE,
        SHEET_CAPABILITY_PROPS, SPEC,
    };

    #[test]
    fn write_only_cell_properties_are_not_advertised_as_reads() {
        let value = CELL_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "value")
            .expect("value capability");
        let clear = CELL_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "clear")
            .expect("clear capability");
        let formula = CELL_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "formula")
            .expect("formula capability");
        // Record 0338: an added cell takes the payload setters; `clear` is
        // set-only.
        assert!(!value.access.get && value.access.set && value.access.add);
        assert!(!clear.access.get && clear.access.set && !clear.access.add);
        assert!(formula.access.get && formula.access.set && formula.access.add);
    }

    #[test]
    fn sheet_write_properties_upgrade_shared_capabilities_only() {
        let name = SHEET_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "name")
            .expect("name capability");
        let hidden = SHEET_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "hidden")
            .expect("hidden capability");
        let visibility = SHEET_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "visibility")
            .expect("visibility capability");
        assert!(name.access.get && name.access.set && name.access.add);
        assert!(hidden.access.get && hidden.access.set && hidden.access.add);
        assert!(visibility.access.get && !visibility.access.set && !visibility.access.add);
    }

    #[test]
    fn merge_registry_preserves_context_read_types_and_scalar_admission() {
        let (_, merge) = MERGE_WRITE_TABLE.resolve("merge").expect("merge setter");
        assert_eq!(merge.spec.ty, crate::value::ValueType::Scalar);
        assert!(merge.clear);
        assert!(matches!(merge.access, crate::registry::Access::Set(_)));

        let cell = CELL_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "merge")
            .expect("cell merge capability");
        assert_eq!(cell.ty, crate::value::ValueType::Text);
        assert!(cell.access.get && cell.access.set && !cell.access.add);

        assert_eq!(RANGE_PROPS.len(), 1);
        assert_eq!(RANGE_PROPS[0].name, "merge");
        assert_eq!(RANGE_PROPS[0].ty, crate::value::ValueType::Bool);
        let range = RANGE_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "merge")
            .expect("range merge capability");
        assert_eq!(range.ty, crate::value::ValueType::Bool);
        assert!(range.access.get && range.access.set && !range.access.add);

        let sheet = SHEET_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "merge")
            .expect("sheet merge capability");
        assert_eq!(sheet.ty, crate::value::ValueType::Scalar);
        assert!(!sheet.access.get && sheet.access.set && !sheet.access.add);

        assert!(!SPEC.gaps.iter().any(|gap| matches!(
            gap.subject,
            crate::node::spec::Subject::Property {
                element: "range",
                key: "merge"
            }
        )));
    }

    #[test]
    fn row_layout_registry_matches_native_setters_and_resets() {
        let (height_index, height) = ROW_WRITE_TABLE.resolve("height").expect("height");
        assert_eq!(height_index, 0);
        assert!(matches!(height.spec.ty, crate::value::ValueType::Length(_)));
        assert!(height.spec.access.get);
        assert!(matches!(height.access, crate::registry::Access::Set(_)));
        assert!(height.clear);

        let (_, hidden) = ROW_WRITE_TABLE.resolve("hidden").expect("hidden");
        assert_eq!(hidden.spec.ty, crate::value::ValueType::Bool);
        assert!(hidden.clear);

        let (outline_index, outline) = ROW_WRITE_TABLE.resolve("outline").expect("outline");
        assert_eq!(outline_index, 2);
        assert!(!outline.spec.access.get);
        assert!(outline.clear);
        assert_eq!(
            ROW_WRITE_TABLE
                .resolve("outlineLevel")
                .expect("outline alias")
                .0,
            outline_index
        );
        assert_eq!(
            ROW_WRITE_TABLE.resolve("group").expect("group alias").0,
            outline_index
        );

        let (_, collapsed) = ROW_WRITE_TABLE.resolve("collapsed").expect("collapsed");
        assert_eq!(collapsed.spec.ty, crate::value::ValueType::Bool);
        assert!(collapsed.clear);

        let read_outline = ROW_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "outlineLevel")
            .expect("read outlineLevel");
        let write_outline = ROW_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "outline")
            .expect("write outline");
        assert!(read_outline.access.get && read_outline.access.set && !read_outline.access.add);
        assert!(!write_outline.access.get && write_outline.access.set);
        for name in ["height", "hidden", "collapsed"] {
            let property = ROW_CAPABILITY_PROPS
                .iter()
                .find(|property| property.name == name)
                .expect("row capability");
            assert!(property.access.get && property.access.set, "{name}");
        }
    }

    #[test]
    fn column_layout_registry_matches_native_setters_and_resets() {
        let (width_index, width) = COLUMN_WRITE_TABLE.resolve("width").expect("width");
        assert_eq!(width_index, 0);
        assert_eq!(width.spec.ty, crate::value::ValueType::Number);
        assert!(matches!(width.access, crate::registry::Access::Set(_)));
        assert!(width.clear);

        let (_, hidden) = COLUMN_WRITE_TABLE.resolve("hidden").expect("hidden");
        assert_eq!(hidden.spec.ty, crate::value::ValueType::Bool);
        assert!(hidden.clear);

        let (outline_index, outline) = COLUMN_WRITE_TABLE
            .resolve("outlineLevel")
            .expect("outline alias");
        assert_eq!(outline_index, 2);
        assert!(!outline.spec.access.get && outline.clear);
        assert_eq!(
            COLUMN_WRITE_TABLE.resolve("group").expect("group alias").0,
            outline_index
        );

        let (_, collapsed) = COLUMN_WRITE_TABLE.resolve("collapsed").expect("collapsed");
        assert_eq!(collapsed.spec.ty, crate::value::ValueType::Bool);
        assert!(collapsed.clear);

        let read_outline = COLUMN_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "outlineLevel")
            .expect("read outlineLevel");
        let write_outline = COLUMN_CAPABILITY_PROPS
            .iter()
            .find(|property| property.name == "outline")
            .expect("write outline");
        assert!(read_outline.access.get && read_outline.access.set && !read_outline.access.add);
        assert!(!write_outline.access.get && write_outline.access.set);
        for name in ["width", "hidden", "collapsed"] {
            let property = COLUMN_CAPABILITY_PROPS
                .iter()
                .find(|property| property.name == name)
                .expect("column capability");
            assert!(property.access.get && property.access.set, "{name}");
        }
    }
}
