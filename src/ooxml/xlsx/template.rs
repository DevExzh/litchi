/// Spreadsheet template module.
///
/// Provides minimal valid templates for creating new Excel workbooks.
/// These templates contain the bare minimum structure required for a valid .xlsx file.
/// Generate a minimal valid workbook.xml content.
///
/// Creates an empty workbook with one default sheet reference.
pub fn default_workbook_xml() -> &'static str {
    include_str!("resources/workbook.xml")
}

/// Generate a minimal valid worksheet.xml content.
///
/// Creates an empty worksheet with default column widths.
pub fn default_worksheet_xml() -> &'static str {
    include_str!("resources/worksheets/sheet1.xml")
}

/// Generate a minimal valid styles.xml content.
///
/// Defines basic cell formats and styles.
pub fn default_styles_xml() -> &'static str {
    include_str!("resources/styles.xml")
}

/// Generate a minimal valid sharedStrings.xml content.
///
/// Creates an empty shared strings table.
pub fn default_shared_strings_xml() -> &'static str {
    include_str!("resources/sharedStrings.xml")
}

/// Generate a minimal valid theme.xml content for Excel.
pub fn default_theme_xml() -> &'static str {
    include_str!("resources/theme/theme1.xml")
}

/// Generate a minimal valid core.xml (core properties) content.
pub fn default_core_props_xml() -> &'static str {
    include_str!("resources/docProps/core.xml")
}

/// Generate a minimal valid app.xml (extended properties) content for Excel.
pub fn default_app_props_xml() -> &'static str {
    include_str!("resources/docProps/app.xml")
}
