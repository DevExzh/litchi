//! Workbook sheet directory metadata from BIFF8 `BoundSheet8` records.

use crate::records::{BoundSheetRecord, SheetType, SheetVisible};

/// Visibility state of a workbook sheet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetVisibility {
    Visible,
    Hidden,
    VeryHidden,
}

/// BIFF substream kind referenced by a workbook sheet entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetKind {
    /// A worksheet or dialog sheet. `BoundSheet8` does not distinguish them.
    WorksheetOrDialog,
    MacroSheet,
    ChartSheet,
    VbaModule,
}

/// One entry in the workbook's sheet directory, in tab order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetMetadata {
    workbook_index: usize,
    name: String,
    visibility: SheetVisibility,
    kind: SheetKind,
    parsed_worksheet_index: Option<usize>,
}

impl SheetMetadata {
    #[must_use]
    pub fn workbook_index(&self) -> usize {
        self.workbook_index
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    #[must_use]
    pub fn visibility(&self) -> SheetVisibility {
        self.visibility
    }
    #[must_use]
    pub fn kind(&self) -> SheetKind {
        self.kind
    }
    /// Index accepted by `Workbook::xls_worksheet`, when this entry was parsed as a worksheet.
    #[must_use]
    pub fn parsed_worksheet_index(&self) -> Option<usize> {
        self.parsed_worksheet_index
    }
    #[must_use]
    pub fn is_visible(&self) -> bool {
        self.visibility == SheetVisibility::Visible
    }

    pub(crate) fn from_bound_sheet(workbook_index: usize, sheet: &BoundSheetRecord) -> Self {
        let visibility = match sheet.visible {
            SheetVisible::Visible => SheetVisibility::Visible,
            SheetVisible::Hidden => SheetVisibility::Hidden,
            SheetVisible::VeryHidden => SheetVisibility::VeryHidden,
        };
        let kind = match sheet.sheet_type {
            SheetType::WorkSheet => SheetKind::WorksheetOrDialog,
            SheetType::MacroSheet => SheetKind::MacroSheet,
            SheetType::ChartSheet => SheetKind::ChartSheet,
            SheetType::VBModule => SheetKind::VbaModule,
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
    use crate::Workbook;
    use std::fs::File;
    use std::path::{Path, PathBuf};

    fn fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    }

    #[test]
    fn reads_hidden_and_very_hidden_poi_sheets() {
        let workbook = Workbook::new(File::open(fixture("45761.xls")).unwrap()).unwrap();
        let sheets = workbook.sheets();
        assert_eq!(sheets.len(), 3);
        assert_eq!(sheets[0].name(), "VisibleSheet");
        assert_eq!(sheets[0].visibility(), SheetVisibility::Visible);
        assert_eq!(sheets[1].name(), "HiddenSheet");
        assert_eq!(sheets[1].visibility(), SheetVisibility::Hidden);
        assert_eq!(sheets[2].name(), "VeryHiddenSheet");
        assert_eq!(sheets[2].visibility(), SheetVisibility::VeryHidden);
        assert!(
            sheets
                .iter()
                .all(|sheet| sheet.kind() == SheetKind::WorksheetOrDialog)
        );

        let hidden = Workbook::new(File::open(fixture("TwoSheetsOneHidden.xls")).unwrap()).unwrap();
        assert_eq!(hidden.sheets()[0].visibility(), SheetVisibility::Hidden);
        assert_eq!(hidden.sheets()[1].visibility(), SheetVisibility::Visible);
    }

    #[test]
    fn catalogs_chart_sheets_without_parsing_them_as_worksheets() {
        let workbook = Workbook::new(File::open(fixture("44010-TwoCharts.xls")).unwrap()).unwrap();
        let sheets = workbook.sheets();
        assert_eq!(sheets.len(), 3);
        assert_eq!(sheets[0].kind(), SheetKind::WorksheetOrDialog);
        assert_eq!(sheets[1].name(), "Graph1");
        assert_eq!(sheets[1].kind(), SheetKind::ChartSheet);
        assert_eq!(sheets[1].parsed_worksheet_index(), None);
        assert_eq!(sheets[2].name(), "Graph2");
        assert_eq!(sheets[2].kind(), SheetKind::ChartSheet);
        assert_eq!(sheets[2].parsed_worksheet_index(), None);
    }
}
