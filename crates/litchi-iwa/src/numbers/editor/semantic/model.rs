//! Public semantic Numbers editor models.

#![allow(unused_imports)]

use super::*;

/// Stable identity and dimensions of a Numbers table.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersTableInfo {
    /// Native model identity retained only inside the IWA adapter.
    pub(crate) object_id: u64,
    /// Checked zero-based position in the editor's semantic table catalog.
    pub index: usize,
    pub name: String,
    pub rows: usize,
    pub columns: usize,
    /// Effective alternating-row and automatic-sizing settings.
    pub appearance: TableAppearance,
    /// Interactive editing lock shown in the Arrange inspector.
    pub lock_state: TableLockState,
}

/// Stable identity and name of a sheet in workbook order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersSheetInfo {
    /// Native identity retained only inside the IWA adapter.
    pub(crate) object_id: u64,
    pub index: usize,
    pub name: String,
}

impl NumbersTableInfo {
    pub(crate) const fn native_id(&self) -> u64 {
        self.object_id
    }
}

impl NumbersSheetInfo {
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
pub struct NumbersTextBoxInfo {
    pub sheet_id: u64,
    pub drawable_object_id: u64,
    pub storage: TextStorageInfo,
}

/// A Numbers text box removed from a sheet with its final text state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RemovedNumbersTextBox {
    pub text_box: NumbersTextBoxInfo,
}

/// A pivot aggregate category that can be used in a formula expression.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NumbersPivotCategoryInfo {
    pub reference: FormulaPivotCategoryReference,
    pub label: Option<String>,
}

/// Horizontal text alignment shared by native Numbers table cells.
pub type NumbersTableCellTextAlignment = TextAlignment;
/// Typed first-line, left, and right indents for a Numbers table cell.
pub type NumbersTableCellParagraphIndents = ParagraphIndents;
/// Typed native line spacing applied to a whole Numbers table cell.
pub type NumbersTableCellParagraphLineSpacing = ParagraphLineSpacing;
/// Typed before/after paragraph spacing applied to a whole Numbers table cell.
pub type NumbersTableCellParagraphSpacing = ParagraphSpacing;
/// Canonical native list preset applied uniformly to a Numbers table cell.
pub type NumbersTableCellParagraphList = ParagraphList;
/// A validated custom text-bullet marker in a Numbers table cell.
pub type NumbersTableCellParagraphListBullet = ParagraphListBullet;
/// Typed marker size and baseline in a Numbers table cell.
pub type NumbersTableCellParagraphListBulletGeometry = ParagraphListBulletGeometry;
/// Typed native list-label and text-gap indentation in a Numbers table cell.
pub type NumbersTableCellParagraphListIndentation = ParagraphListIndentation;
pub type NumbersTableCellParagraphListLabelColor = ParagraphListLabelColor;
/// Locale-aware numbered-list label format in a Numbers table cell.
pub type NumbersTableCellParagraphListNumberFormat = ParagraphListNumberFormat;
/// Flat or hierarchical numbered-list labels in a Numbers table cell.
pub type NumbersTableCellParagraphListNumberTiering = ParagraphListNumberTiering;
/// Number-label size for a numbered paragraph in a Numbers table cell.
pub type NumbersTableCellParagraphListNumberScale = ParagraphListNumberScale;
/// A validated zero-based list nesting level in a Numbers table cell.
pub type NumbersTableCellParagraphListLevel = ParagraphListLevel;
/// One effective list-level boundary in a Numbers table cell.
pub type NumbersTableCellParagraphListLevelPlacement = ParagraphListLevelPlacement;
/// Whether a Numbers table-cell paragraph continues or restarts list numbering.
pub type NumbersTableCellParagraphListNumbering = ParagraphListNumbering;
/// One paragraph-scoped list preset boundary in a Numbers table cell.
pub type NumbersTableCellParagraphListPlacement = ParagraphListPlacement;
/// Ordered explicit ruler tab stops for a Numbers table cell.
pub type NumbersTableCellParagraphTabStops = ParagraphTabStops;
/// Typed solid background painted behind a whole Numbers table cell's text.
pub type NumbersTableCellTextBackground = TextBackground;
/// Validated custom baseline displacement applied to a whole Numbers cell.
pub type NumbersTableCellTextBaselineShift = TextBaselineShift;
/// Typed capitalization applied to a whole Numbers table cell.
pub type NumbersTableCellTextCapitalization = TextCapitalization;
/// Validated tracking applied to a whole Numbers table cell.
pub type NumbersTableCellTextCharacterSpacing = TextCharacterSpacing;
/// Validated foreground color applied to a whole Numbers table cell.
pub type NumbersTableCellTextColor = RgbaColor;
/// Typed underline and strikethrough formatting for a whole Numbers table cell.
pub type NumbersTableCellTextDecorations = TextDecorations;
/// Strict PostScript font identity applied to a whole Numbers table cell.
pub type NumbersTableCellTextFont = TextFont;
/// Typed ligature policy applied to a whole Numbers table cell.
pub type NumbersTableCellTextLigatures = TextLigatures;
/// Typed outline applied to a whole Numbers table cell's text.
pub type NumbersTableCellTextOutline = TextOutline;
/// Typed normal, superscript, or subscript formatting for a whole Numbers cell.
pub type NumbersTableCellTextScript = TextScript;
/// Typed drop shadow applied to a whole Numbers table cell's text.
pub type NumbersTableCellTextShadow = TextShadow;
/// Whole-cell point size, bold, and italic formatting.
pub type NumbersTableCellTextStyle = TextStyle;

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
pub struct NumbersEditor {
    pub(in crate::numbers::editor) package: IWorkPackage,
}
