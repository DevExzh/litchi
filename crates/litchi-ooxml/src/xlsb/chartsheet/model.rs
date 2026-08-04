//! Typed inert model for the XLSB Chart Sheet stream (MS-XLSB 2.1.7.7).
//!
//! A chart sheet is a sheet that contains a single chart. Per MS-XLSB
//! 2.1.7.5/2.1.7.6/2.1.7.23 the chart itself, its drawing, and the sheet's
//! drawing are standard DrawingML XML parts (identical to XLSX); only the
//! chart sheet part is a BIFF12 binary stream. The records of that stream
//! expose sheet-level metadata only (code name, publish flag, tab color,
//! views, protection, page setup, and drawing-part links) — the chart type
//! and plot definition live in the linked `c:chartSpace` XML part and are
//! surfaced through [`crate::xlsb::drawing::SheetDrawing`].
//!
//! All values are inert data snapshots: relationship identifiers, password
//! verifiers, and hash data are stored verbatim and are never dereferenced,
//! verified, or executed.

use crate::xlsb::worksheet::StrongProtection;

/// Visibility state of a chart sheet, from `BrtBundleSh hsState`
/// (MS-XLSB 2.4.718).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbChartSheetState {
    /// The sheet is visible.
    Visible,
    /// The sheet is hidden through the user interface.
    Hidden,
    /// The sheet is hidden and cannot be shown through the user interface.
    VeryHidden,
}

/// `BrtColor` payload as carried by `BrtCsProp brtcolorTab` (MS-XLSB
/// 2.4.337): the background color of the sheet tab.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbChartSheetColor {
    /// Whether the palette color matches the RGB values (`fValidRGB`).
    pub valid_rgb: bool,
    /// Kind of color information carried by the record (`xColorType`).
    pub color_type: XlsbChartSheetColorType,
    /// Palette or theme index (`index`); meaningless for RGB colors.
    pub index: u8,
    /// Tint and shade value (`nTintAndShade`).
    pub tint: i16,
    /// Red, green, blue, and alpha components.
    pub rgba: [u8; 4],
}

impl XlsbChartSheetColor {
    /// The default tab color when `BrtCsProp` is absent: automatic.
    pub fn automatic() -> Self {
        XlsbChartSheetColor {
            valid_rgb: false,
            color_type: XlsbChartSheetColorType::Automatic,
            index: 0,
            tint: 0,
            rgba: [0; 4],
        }
    }
}

impl Default for XlsbChartSheetColor {
    fn default() -> Self {
        Self::automatic()
    }
}

/// `xColorType` values of a `BrtColor` (MS-XLSB 2.4.337).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsbChartSheetColorType {
    /// Color information is automatically determined by the application.
    Automatic,
    /// A color from the color palette, addressed by `index` (an `Icv`).
    Indexed,
    /// A standard ARGB color carried by the RGBA components.
    Rgb,
    /// A theme color addressed by `index`.
    Theme,
}

/// One chart sheet view from `BrtBeginCsView` (MS-XLSB 2.4.38).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbChartSheetView {
    /// Whether the chart sheet is currently selected (`fSelected`).
    pub selected: bool,
    /// Window zoom level as a percentage; 0 means no zoom level is set
    /// (`wScale`, 10..=400 or 0).
    pub scale: u32,
    /// Zero-based index of the associated `BrtBookView` in the workbook part
    /// (`iwbkview`); stored verbatim and never dereferenced.
    pub workbook_view_index: u32,
}

/// Chart sheet protection from `BrtCsProtection` (MS-XLSB 2.4.345).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct XlsbChartSheetProtection {
    /// Password verifier value; 0 means no password is required (`protpwd`).
    /// Stored verbatim and never verified.
    pub password_verifier: u16,
    /// Whether changes to the chart elements are prevented (`fLocked`).
    pub locked: bool,
    /// Whether changes to graphic objects are prevented (`fObjects`).
    pub objects: bool,
}

/// Page layout and printing settings from `BrtCsPageSetup` (MS-XLSB 2.4.343).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbChartSheetPageSetup {
    /// Printer paper size (`iPaperSize`; 1..=118 are standard sizes).
    pub paper_size: u32,
    /// Horizontal printer resolution in dots per inch (`iRes`).
    pub horizontal_resolution: u32,
    /// Vertical printer resolution in dots per inch (`iVRes`).
    pub vertical_resolution: u32,
    /// Number of copies to print (`iCopies`).
    pub copies: u32,
    /// Starting page number when `use_page_start` is set (`iPageStart`).
    pub page_start: i16,
    /// Landscape orientation when `use_default_orientation` is clear
    /// (`fLandscape`).
    pub landscape: bool,
    /// Print in black and white (`fNoColor`).
    pub black_and_white: bool,
    /// Orientation is determined by the application and printer (`fNoOrient`).
    pub use_default_orientation: bool,
    /// `page_start` is used as the first page number (`fUsePage`).
    pub use_page_start: bool,
    /// Graphics are omitted from the printed page (`fDraft`).
    pub draft: bool,
    /// Relationship identifier of the Printer Settings part (`szRelID`);
    /// stored verbatim and never dereferenced.
    pub printer_settings_rel_id: String,
}

/// Typed inert model of one Chart Sheet part (MS-XLSB 2.1.7.7).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsbChartSheet {
    /// Sheet name from `BrtBundleSh` in the workbook part.
    pub name: String,
    /// Sheet visibility state from `BrtBundleSh`.
    pub state: XlsbChartSheetState,
    /// Code name from `BrtCsProp strName` (MS-XLSB 2.4.344); empty when the
    /// record is absent.
    pub code_name: String,
    /// Whether the chart sheet is published (`BrtCsProp fPublish`).
    pub published: bool,
    /// Sheet tab background color (`BrtCsProp brtcolorTab`).
    pub tab_color: XlsbChartSheetColor,
    /// Chart sheet views from the `BrtBeginCsViews` collection.
    pub views: Vec<XlsbChartSheetView>,
    /// Classic chart sheet protection, when present.
    pub protection: Option<XlsbChartSheetProtection>,
    /// ISO strong password data from `BrtCsProtectionIso`, when present.
    pub strong_protection: Option<StrongProtection>,
    /// Page layout and printing settings, when present.
    pub page_setup: Option<XlsbChartSheetPageSetup>,
    /// Relationship identifier of the Drawings part from `BrtDrawing`
    /// (MS-XLSB 2.4.354); stored verbatim and never dereferenced.
    pub drawing_rel_id: Option<String>,
    /// Relationship identifier of the VML Drawings part from
    /// `BrtLegacyDrawing` (MS-XLSB 2.4.703).
    pub legacy_drawing_rel_id: Option<String>,
    /// Relationship identifier of the header/footer VML Drawings part from
    /// `BrtLegacyDrawingHF` (MS-XLSB 2.4.704).
    pub legacy_drawing_header_footer_rel_id: Option<String>,
}
