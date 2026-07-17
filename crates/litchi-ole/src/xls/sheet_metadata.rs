//! Workbook sheet directory metadata from BIFF8 `BoundSheet8` records.

use crate::xls::records::{BoundSheetRecord, SheetType, SheetVisible};

/// Visibility state of a workbook sheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

/// BIFF substream kind referenced by a workbook sheet entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsSheetKind {
    /// A worksheet or dialog sheet. `BoundSheet8` does not distinguish them.
    WorksheetOrDialog,
    MacroSheet,
    ChartSheet,
    VbaModule,
}

/// One entry in the workbook's sheet directory, in tab order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XlsSheetMetadata {
    workbook_index: usize,
    name: String,
    visibility: XlsSheetVisibility,
    kind: XlsSheetKind,
    parsed_worksheet_index: Option<usize>,
}

impl XlsSheetMetadata {
    pub fn workbook_index(&self) -> usize {
        self.workbook_index
    }
    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn visibility(&self) -> XlsSheetVisibility {
        self.visibility
    }
    pub fn kind(&self) -> XlsSheetKind {
        self.kind
    }
    /// Index accepted by `XlsWorkbook::xls_worksheet`, when this entry was parsed as a worksheet.
    pub fn parsed_worksheet_index(&self) -> Option<usize> {
        self.parsed_worksheet_index
    }
    pub fn is_visible(&self) -> bool {
        self.visibility == XlsSheetVisibility::Visible
    }

    pub(crate) fn from_bound_sheet(workbook_index: usize, sheet: &BoundSheetRecord) -> Self {
        let visibility = match sheet.visible {
            SheetVisible::Visible => XlsSheetVisibility::Visible,
            SheetVisible::Hidden => XlsSheetVisibility::Hidden,
            SheetVisible::VeryHidden => XlsSheetVisibility::VeryHidden,
        };
        let kind = match sheet.sheet_type {
            SheetType::WorkSheet => XlsSheetKind::WorksheetOrDialog,
            SheetType::MacroSheet => XlsSheetKind::MacroSheet,
            SheetType::ChartSheet => XlsSheetKind::ChartSheet,
            SheetType::VBModule => XlsSheetKind::VbaModule,
        };
        Self {
            workbook_index,
            name: sheet.name.clone(),
            visibility,
            kind,
            parsed_worksheet_index: None,
        }
    }

    pub(crate) fn set_parsed_worksheet_index(&mut self, index: usize) {
        self.parsed_worksheet_index = Some(index);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::XlsWorkbook;
    use std::fs::File;
    use std::path::{Path, PathBuf};

    fn fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/poi/test-data/spreadsheet")
            .join(name)
    }

    #[test]
    fn reads_hidden_and_very_hidden_poi_sheets() {
        let workbook = XlsWorkbook::new(File::open(fixture("45761.xls")).unwrap()).unwrap();
        let sheets = workbook.sheets();
        assert_eq!(sheets.len(), 3);
        assert_eq!(sheets[0].name(), "VisibleSheet");
        assert_eq!(sheets[0].visibility(), XlsSheetVisibility::Visible);
        assert_eq!(sheets[1].name(), "HiddenSheet");
        assert_eq!(sheets[1].visibility(), XlsSheetVisibility::Hidden);
        assert_eq!(sheets[2].name(), "VeryHiddenSheet");
        assert_eq!(sheets[2].visibility(), XlsSheetVisibility::VeryHidden);
        assert!(
            sheets
                .iter()
                .all(|sheet| sheet.kind() == XlsSheetKind::WorksheetOrDialog)
        );

        let hidden =
            XlsWorkbook::new(File::open(fixture("TwoSheetsOneHidden.xls")).unwrap()).unwrap();
        assert_eq!(hidden.sheets()[0].visibility(), XlsSheetVisibility::Hidden);
        assert_eq!(hidden.sheets()[1].visibility(), XlsSheetVisibility::Visible);
    }

    #[test]
    fn catalogs_chart_sheets_without_parsing_them_as_worksheets() {
        let workbook =
            XlsWorkbook::new(File::open(fixture("44010-TwoCharts.xls")).unwrap()).unwrap();
        let sheets = workbook.sheets();
        assert_eq!(sheets.len(), 3);
        assert_eq!(sheets[0].kind(), XlsSheetKind::WorksheetOrDialog);
        assert_eq!(sheets[1].name(), "Graph1");
        assert_eq!(sheets[1].kind(), XlsSheetKind::ChartSheet);
        assert_eq!(sheets[1].parsed_worksheet_index(), None);
        assert_eq!(sheets[2].name(), "Graph2");
        assert_eq!(sheets[2].kind(), XlsSheetKind::ChartSheet);
        assert_eq!(sheets[2].parsed_worksheet_index(), None);
    }
}
