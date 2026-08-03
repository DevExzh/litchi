//! Immutable XLSX worksheet sheet-format-properties read model.

use crate::error::{OoxmlError, Result};

/// Effective worksheet defaults and outline metadata from `sheetFormatPr`.
///
/// This remains a compatibility view for the legacy OOXML host API. The
/// checked semantic model and all XML validation live in `litchi-xlsx`.
#[derive(Debug, Clone, PartialEq)]
pub struct WorksheetSheetFormatProperties {
    base_column_width: u32,
    default_column_width: Option<f64>,
    default_row_height: f64,
    custom_height: bool,
    zero_height: bool,
    thick_top: bool,
    thick_bottom: bool,
    outline_level_row: u8,
    outline_level_column: u8,
    dy_descent: Option<f64>,
}

impl WorksheetSheetFormatProperties {
    pub fn base_column_width(&self) -> u32 {
        self.base_column_width
    }

    pub fn default_column_width(&self) -> Option<f64> {
        self.default_column_width
    }

    pub fn effective_default_column_width(&self) -> f64 {
        self.default_column_width
            .unwrap_or(self.base_column_width as f64)
    }

    pub fn default_row_height(&self) -> f64 {
        self.default_row_height
    }

    pub fn custom_height(&self) -> bool {
        self.custom_height
    }

    pub fn zero_height(&self) -> bool {
        self.zero_height
    }

    pub fn thick_top(&self) -> bool {
        self.thick_top
    }

    pub fn thick_bottom(&self) -> bool {
        self.thick_bottom
    }

    pub fn outline_level_row(&self) -> u8 {
        self.outline_level_row
    }

    pub fn outline_level_column(&self) -> u8 {
        self.outline_level_column
    }

    /// Excel 2010 typographical descent in pixels at 100% worksheet zoom.
    pub fn dy_descent(&self) -> Option<f64> {
        self.dy_descent
    }
}

impl From<&litchi_xlsx::layout::Defaults> for WorksheetSheetFormatProperties {
    fn from(defaults: &litchi_xlsx::layout::Defaults) -> Self {
        Self {
            base_column_width: u32::from(defaults.base_width()),
            default_column_width: defaults.width().map(litchi_xlsx::layout::Width::get),
            default_row_height: defaults.height().get(),
            custom_height: defaults.custom_height(),
            zero_height: defaults.hidden(),
            thick_top: defaults.thick_top(),
            thick_bottom: defaults.thick_bottom(),
            outline_level_row: defaults.row_outline().get(),
            outline_level_column: defaults.column_outline().get(),
            dy_descent: defaults.descent().map(litchi_xlsx::layout::Descent::get),
        }
    }
}

/// Parse the worksheet's direct `sheetFormatPr` child through the canonical
/// XLSX owner.
pub fn parse_worksheet_sheet_format_properties(
    xml: &[u8],
) -> Result<Option<WorksheetSheetFormatProperties>> {
    let defaults = litchi_xlsx::raw::parse_worksheet_defaults(xml).map_err(map_xlsx_error)?;
    Ok(defaults.as_ref().map(WorksheetSheetFormatProperties::from))
}

fn map_xlsx_error(error: litchi_xlsx::Error) -> OoxmlError {
    match error {
        litchi_xlsx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_xlsx::Error::MarkupCompatibility(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Xml(error) => OoxmlError::from(error),
        litchi_xlsx::Error::Common(error) => OoxmlError::Common(error),
        other => OoxmlError::Xlsx(other),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetSheetFormatProperties>> {
        parse_worksheet_sheet_format_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_all_core_attributes_and_effective_defaults() {
        let value = parse(concat!(
            r#"<sheetFormatPr baseColWidth="9" defaultColWidth="11.5" "#,
            r#"defaultRowHeight="18.25" customHeight="false" zeroHeight="1" "#,
            r#"thickTop="true" thickBottom="1" outlineLevelRow="4" outlineLevelCol="3"/>"#,
        ))
        .unwrap()
        .unwrap();
        assert_eq!(value.base_column_width(), 9);
        assert_eq!(value.default_column_width(), Some(11.5));
        assert_eq!(value.effective_default_column_width(), 11.5);
        assert_eq!(value.default_row_height(), 18.25);
        assert!(!value.custom_height());
        assert!(value.zero_height() && value.thick_top() && value.thick_bottom());
        assert_eq!(value.outline_level_row(), 4);
        assert_eq!(value.outline_level_column(), 3);

        let defaults = parse(r#"<sheetFormatPr defaultRowHeight="15"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(defaults.base_column_width(), 8);
        assert_eq!(defaults.default_column_width(), None);
        assert_eq!(defaults.effective_default_column_width(), 8.0);
        assert!(!defaults.zero_height());
        assert_eq!(defaults.outline_level_row(), 0);
    }

    #[test]
    fn supports_strict_namespace_and_direct_child_only() {
        let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetFormatPr defaultRowHeight="16"/></worksheet>"#;
        assert_eq!(
            parse_worksheet_sheet_format_properties(strict)
                .unwrap()
                .unwrap()
                .default_row_height(),
            16.0
        );
        let nested = format!(
            r#"<worksheet xmlns="{NS}"><wrapper><sheetFormatPr defaultRowHeight="15"/></wrapper></worksheet>"#
        );
        assert!(
            parse_worksheet_sheet_format_properties(nested.as_bytes())
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn dy_descent_survives_mce_and_forces_custom_height() {
        let xml = format!(
            concat!(
                r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
                r#"xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">"#,
                r#"<sheetFormatPr defaultRowHeight="15" customHeight="0" x14ac:dyDescent="0.25"/></worksheet>"#,
            ),
            NS
        );
        let value = parse_worksheet_sheet_format_properties(xml.as_bytes())
            .unwrap()
            .unwrap();
        assert_eq!(value.dy_descent(), Some(0.25));
        assert!(value.custom_height());
    }

    #[test]
    fn rejects_invalid_bounds_attributes_and_leaf_content() {
        for child in [
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelRow="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" outlineLevelCol="8"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" baseColWidth="256"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" defaultColWidth="65536"/>"#,
            r#"<sheetFormatPr defaultRowHeight="NaN"/>"#,
            r#"<sheetFormatPr/>"#,
            r#"<sheetFormatPr defaultRowHeight="15" mystery="1"/>"#,
            r#"<sheetFormatPr defaultRowHeight="15"><child/></sheetFormatPr>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
        assert!(
            parse(concat!(
                r#"<sheetFormatPr defaultRowHeight="15"/>"#,
                r#"<sheetFormatPr defaultRowHeight="15"/>"#
            ))
            .is_err()
        );
    }

    fn fixture_sheet(bytes: &[u8]) -> WorksheetSheetFormatProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        parse_worksheet_sheet_format_properties(part.blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn reads_libreoffice_hidden_default_rows_fixture() {
        let value = fixture_sheet(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf105840_allRowsHidden.xlsx"
        )));
        assert_eq!(value.default_row_height(), 15.0);
        assert!(value.zero_height());
    }

    #[test]
    fn reads_libreoffice_custom_dimensions_fixture() {
        let value = fixture_sheet(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf120168.xlsx"
        )));
        assert_eq!(value.default_column_width(), Some(21.85546875));
        assert_eq!(value.default_row_height(), 39.0);
        assert_eq!(value.dy_descent(), Some(0.25));
        assert!(value.custom_height());
    }
}
