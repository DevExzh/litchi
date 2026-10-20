//! Public semantic Numbers editor models.

#![allow(unused_imports)]

use super::*;

/// Stable identity and dimensions of a Numbers table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TableInfo {
    /// Native model identity retained only inside the IWA adapter.
    pub(crate) object_id: u64,
    /// Checked zero-based position in the editor's semantic table catalog.
    pub index: usize,
    pub name: String,
    pub rows: usize,
    pub columns: usize,
    /// Effective alternating-row and automatic-sizing settings.
    pub appearance: Appearance,
}

/// Stable identity and name of a sheet in workbook order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetInfo {
    /// Native identity retained only inside the IWA adapter.
    pub(crate) object_id: u64,
    pub index: usize,
    pub name: String,
}

impl TableInfo {
    /// Return the stable native table identifier accepted by cell-level APIs.
    #[must_use]
    pub const fn id(&self) -> u64 {
        self.object_id
    }

    pub(crate) const fn native_id(&self) -> u64 {
        self.object_id
    }
}

impl SheetInfo {
    /// Return the stable native sheet identifier accepted by sheet-scoped APIs.
    #[must_use]
    pub const fn id(&self) -> u64 {
        self.object_id
    }

    pub(crate) const fn native_id(&self) -> u64 {
        self.object_id
    }
}

/// A writable ordinary text box owned by one Numbers sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextBoxInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// A Numbers text box removed from a sheet with its final text state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedNumbersTextBox {
    pub text_box: TextBoxInfo,
}

/// A pivot aggregate category that can be used in a formula expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PivotCategoryInfo {
    pub reference: FormulaPivotCategoryReference,
    pub label: Option<String>,
}

/// Horizontal text alignment shared by native Numbers table cells.
pub type TableCellTextAlignment = Alignment;
/// Typed first-line, left, and right indents for a Numbers table cell.
pub type TableCellParagraphIndents = Indents;
/// Typed native line spacing applied to a whole Numbers table cell.
pub type TableCellParagraphLineSpacing = LineSpacing;
/// Typed before/after paragraph spacing applied to a whole Numbers table cell.
pub type TableCellParagraphSpacing = Spacing;
/// Canonical native list preset applied uniformly to a Numbers table cell.
pub type TableCellParagraphList = ParagraphList;
/// A validated custom text-bullet marker in a Numbers table cell.
pub type TableCellParagraphListBullet = ParagraphListBullet;
/// Typed marker size and baseline in a Numbers table cell.
pub type TableCellParagraphListBulletGeometry = ParagraphListBulletGeometry;
/// Typed native list-label and text-gap indentation in a Numbers table cell.
pub type TableCellParagraphListIndentation = ParagraphListIndentation;
pub type TableCellParagraphListLabelColor = ParagraphListLabelColor;
/// Locale-aware numbered-list label format in a Numbers table cell.
pub type TableCellParagraphListNumberFormat = ParagraphListNumberFormat;
/// Flat or hierarchical numbered-list labels in a Numbers table cell.
pub type TableCellParagraphListNumberTiering = ParagraphListNumberTiering;
/// Number-label size for a numbered paragraph in a Numbers table cell.
pub type TableCellParagraphListNumberScale = ParagraphListNumberScale;
/// A validated zero-based list nesting level in a Numbers table cell.
pub type TableCellParagraphListLevel = ParagraphListLevel;
/// One effective list-level boundary in a Numbers table cell.
pub type TableCellParagraphListLevelPlacement = ParagraphListLevelPlacement;
/// Whether a Numbers table-cell paragraph continues or restarts list numbering.
pub type TableCellParagraphListNumbering = ParagraphListNumbering;
/// One paragraph-scoped list preset boundary in a Numbers table cell.
pub type TableCellParagraphListPlacement = ParagraphListPlacement;
/// Ordered explicit ruler tab stops for a Numbers table cell.
pub type TableCellParagraphTabStops = ParagraphTabStops;
/// Typed solid background painted behind a whole Numbers table cell's text.
pub type TableCellTextBackground = Background;
/// Validated custom baseline displacement applied to a whole Numbers cell.
pub type TableCellTextBaselineShift = TextBaselineShift;
/// Typed capitalization applied to a whole Numbers table cell.
pub type TableCellTextCapitalization = TextCapitalization;
/// Validated tracking applied to a whole Numbers table cell.
pub type TableCellTextCharacterSpacing = TextCharacterSpacing;
/// Validated foreground color applied to a whole Numbers table cell.
pub type TableCellTextColor = RgbaColor;
/// Typed underline and strikethrough formatting for a whole Numbers table cell.
pub type TableCellTextDecorations = TextDecorations;
/// Strict PostScript font identity applied to a whole Numbers table cell.
pub type TableCellTextFont = TextFont;
/// Typed ligature policy applied to a whole Numbers table cell.
pub type TableCellTextLigatures = TextLigatures;
/// Typed outline applied to a whole Numbers table cell's text.
pub type TableCellTextOutline = Outline;
/// Typed normal, superscript, or subscript formatting for a whole Numbers cell.
pub type TableCellTextScript = TextScript;
/// Typed drop shadow applied to a whole Numbers table cell's text.
pub type TableCellTextShadow = Shadow;
/// Whole-cell point size, bold, and italic formatting.
pub type TableCellTextStyle = TextStyle;

/// Storage identity and rule count of conditional highlighting attached to one cell.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableCellConditionalHighlightInfo {
    pub table_id: u64,
    pub row: usize,
    pub column: usize,
    pub list_identifier: u32,
    pub style_set_object_id: u64,
    pub rule_count: u32,
}

/// Mutable, transactional Numbers package editor.
///
/// Each semantic edit is applied to a cloned package and committed only after
/// all affected IWA components serialize successfully.
#[derive(Debug, Clone)]
pub struct Draft {
    pub(in crate::numbers::editor) package: HostPackage,
}
