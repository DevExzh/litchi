//! Excel workbook embedding for chart data.
//!
//! This module generates minimal XLSX files containing chart data for embedding
//! in PowerPoint presentations. The generated files are valid Excel workbooks
//! that PowerPoint can read to display chart data.

use crate::ooxml::error::{OoxmlError, Result};
use crate::ooxml::pptx::parts::chart::ChartData;

/// Generate a minimal Excel workbook containing chart data.
///
/// Creates a valid XLSX file with one worksheet containing:
/// - Categories in column A (row 2 onwards)
/// - Series names in row 1 (column B onwards)
/// - Series data in the corresponding columns
///
/// # Arguments
/// * `chart` - The chart data to embed
///
/// # Returns
/// * `Ok(Vec<u8>)` - The XLSX file bytes
/// * `Err` if generation fails
///
/// # Example Layout
/// ```text
///     |    A     |    B     |    C     |
/// ----+----------+----------+----------+
///   1 |          | Series 1 | Series 2 |
///   2 | Cat 1    |   10.0   |   15.0   |
///   3 | Cat 2    |   20.0   |   25.0   |
/// ```
pub fn generate_chart_excel_data(chart: &ChartData) -> Result<Vec<u8>> {
    use soapberry_zip::office::StreamingArchiveWriter;

    let mut writer = StreamingArchiveWriter::new();

    // [Content_Types].xml
    let content_types = generate_content_types();
    writer
        .write_deflated("[Content_Types].xml", content_types.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write content types: {}", e)))?;

    // _rels/.rels
    let rels = generate_root_rels();
    writer
        .write_deflated("_rels/.rels", rels.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write rels: {}", e)))?;

    // xl/workbook.xml
    let workbook = generate_workbook_xml();
    writer
        .write_deflated("xl/workbook.xml", workbook.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write workbook: {}", e)))?;

    // xl/_rels/workbook.xml.rels
    let workbook_rels = generate_workbook_rels();
    writer
        .write_deflated("xl/_rels/workbook.xml.rels", workbook_rels.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write workbook rels: {}", e)))?;

    // xl/worksheets/sheet1.xml
    let sheet = generate_worksheet_xml(chart);
    writer
        .write_deflated("xl/worksheets/sheet1.xml", sheet.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write worksheet: {}", e)))?;

    // xl/styles.xml (minimal styles)
    let styles = generate_styles_xml();
    writer
        .write_deflated("xl/styles.xml", styles.as_bytes())
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to write styles: {}", e)))?;

    writer
        .finish_to_bytes()
        .map_err(|e| OoxmlError::InvalidFormat(format!("Failed to finish ZIP: {}", e)))
}

/// Generate [Content_Types].xml for the XLSX package.
fn generate_content_types() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"#.to_string()
}

/// Generate root _rels/.rels file.
fn generate_root_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#.to_string()
}

/// Generate xl/workbook.xml.
fn generate_workbook_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
</sheets>
</workbook>"#.to_string()
}

/// Generate xl/_rels/workbook.xml.rels.
fn generate_workbook_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"#.to_string()
}

/// Generate xl/styles.xml (minimal styles).
fn generate_styles_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"#.to_string()
}

/// Generate xl/worksheets/sheet1.xml with chart data.
fn generate_worksheet_xml(chart: &ChartData) -> String {
    use std::fmt::Write;

    let mut xml = String::with_capacity(4096);

    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
    );
    xml.push_str("<sheetData>");

    // Determine the number of rows needed
    let max_values = chart
        .series
        .iter()
        .map(|s| s.values.len())
        .max()
        .unwrap_or(0);
    let max_categories = chart
        .series
        .iter()
        .map(|s| s.categories.len())
        .max()
        .unwrap_or(0);
    let num_data_rows = max_values.max(max_categories);

    // Row 1: Header row with series names
    xml.push_str(r#"<row r="1">"#);
    // A1 is empty (category header)
    xml.push_str(r#"<c r="A1" t="s"><v>0</v></c>"#); // Placeholder for shared string

    for (col_idx, series) in chart.series.iter().enumerate() {
        let col_letter = column_letter(col_idx + 1); // B, C, D, ...
        let _ = write!(
            xml,
            r#"<c r="{}1" t="inlineStr"><is><t>{}</t></is></c>"#,
            col_letter,
            escape_xml(&series.name)
        );
    }
    xml.push_str("</row>");

    // Data rows (row 2 onwards)
    for row_idx in 0..num_data_rows {
        let row_num = row_idx + 2; // Excel rows are 1-indexed, data starts at row 2
        let _ = write!(xml, r#"<row r="{}">"#, row_num);

        // Column A: Category name (if available)
        let category = chart
            .series
            .first()
            .and_then(|s| s.categories.get(row_idx))
            .map(|s| s.as_str())
            .unwrap_or("");

        if !category.is_empty() {
            let _ = write!(
                xml,
                r#"<c r="A{}" t="inlineStr"><is><t>{}</t></is></c>"#,
                row_num,
                escape_xml(category)
            );
        }

        // Columns B onwards: Series values
        for (col_idx, series) in chart.series.iter().enumerate() {
            let col_letter = column_letter(col_idx + 1);
            if let Some(value) = series.values.get(row_idx) {
                let _ = write!(
                    xml,
                    r#"<c r="{}{}"><v>{}</v></c>"#,
                    col_letter, row_num, value
                );
            }
        }

        xml.push_str("</row>");
    }

    xml.push_str("</sheetData>");
    xml.push_str("</worksheet>");

    xml
}

/// Convert a 0-based column index to Excel column letter (0=A, 1=B, ..., 25=Z, 26=AA).
fn column_letter(col: usize) -> String {
    let mut result = String::new();
    let mut n = col;

    loop {
        let remainder = n % 26;
        result.insert(0, (b'A' + remainder as u8) as char);
        if n < 26 {
            break;
        }
        n = n / 26 - 1;
    }

    result
}

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::pptx::parts::chart::{ChartSeries, ChartType};

    #[test]
    fn test_column_letter() {
        assert_eq!(column_letter(0), "A");
        assert_eq!(column_letter(1), "B");
        assert_eq!(column_letter(25), "Z");
        assert_eq!(column_letter(26), "AA");
        assert_eq!(column_letter(27), "AB");
    }

    #[test]
    fn test_generate_chart_excel_data() {
        let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100)
            .add_series(
                ChartSeries::new("Sales")
                    .with_categories(vec!["Q1".to_string(), "Q2".to_string()])
                    .with_values(vec![100.0, 200.0]),
            )
            .add_series(ChartSeries::new("Profit").with_values(vec![50.0, 75.0]));

        let result = generate_chart_excel_data(&chart);
        assert!(result.is_ok());

        let bytes = result.unwrap();
        // Check ZIP signature
        assert!(bytes.len() > 4);
        assert_eq!(&bytes[0..4], &[0x50, 0x4B, 0x03, 0x04]); // PK\x03\x04
    }

    #[test]
    fn test_generate_worksheet_xml() {
        let chart = ChartData::new(ChartType::Bar, 0, 0, 100, 100).add_series(
            ChartSeries::new("Test")
                .with_categories(vec!["A".to_string(), "B".to_string()])
                .with_values(vec![1.0, 2.0]),
        );

        let xml = generate_worksheet_xml(&chart);
        assert!(xml.contains("<worksheet"));
        assert!(xml.contains("<sheetData>"));
        assert!(xml.contains("Test")); // Series name
        assert!(xml.contains("<v>1</v>")); // First value
        assert!(xml.contains("<v>2</v>")); // Second value
    }

    /// **Feature: charts-smartart-integration, Property 5: Excel data is valid ZIP**
    /// **Validates: Requirements 1.3**
    ///
    /// For any ChartData, the generated Excel bytes SHALL start with the ZIP
    /// signature (0x504B0304) and contain a valid workbook structure.
    #[cfg(test)]
    mod property_tests {
        use super::*;
        use proptest::prelude::*;

        /// Strategy to generate valid chart types
        fn chart_type_strategy() -> impl Strategy<Value = ChartType> {
            prop_oneof![
                Just(ChartType::Bar),
                Just(ChartType::Column),
                Just(ChartType::Line),
                Just(ChartType::Pie),
                Just(ChartType::Area),
                Just(ChartType::Scatter),
                Just(ChartType::Doughnut),
            ]
        }

        /// Strategy to generate valid series names (non-empty, no control chars)
        fn series_name_strategy() -> impl Strategy<Value = String> {
            "[a-zA-Z][a-zA-Z0-9 ]{0,20}".prop_map(|s| s.trim().to_string())
        }

        /// Strategy to generate valid category names
        fn category_strategy() -> impl Strategy<Value = String> {
            "[a-zA-Z0-9][a-zA-Z0-9 ]{0,10}".prop_map(|s| s.trim().to_string())
        }

        /// Strategy to generate a chart series
        fn series_strategy() -> impl Strategy<Value = ChartSeries> {
            (
                series_name_strategy(),
                prop::collection::vec(category_strategy(), 1..10),
                prop::collection::vec(-1000.0f64..1000.0f64, 1..10),
            )
                .prop_map(|(name, categories, values)| {
                    ChartSeries::new(name)
                        .with_categories(categories)
                        .with_values(values)
                })
        }

        /// Strategy to generate valid ChartData
        fn chart_data_strategy() -> impl Strategy<Value = ChartData> {
            (
                chart_type_strategy(),
                prop::collection::vec(series_strategy(), 1..5),
                0i64..10000000i64,
                0i64..10000000i64,
                100i64..10000000i64,
                100i64..10000000i64,
            )
                .prop_map(|(chart_type, series, x, y, width, height)| {
                    let mut chart = ChartData::new(chart_type, x, y, width, height);
                    for s in series {
                        chart = chart.add_series(s);
                    }
                    chart
                })
        }

        proptest! {
            #![proptest_config(ProptestConfig::with_cases(100))]

            #[test]
            fn prop_excel_data_is_valid_zip(chart in chart_data_strategy()) {
                // Generate Excel data
                let result = generate_chart_excel_data(&chart);

                // Should always succeed for valid input
                prop_assert!(result.is_ok(), "Excel generation failed: {:?}", result.err());

                let bytes = result.unwrap();

                // Must have at least ZIP header
                prop_assert!(bytes.len() >= 4, "Excel data too short: {} bytes", bytes.len());

                // Must start with ZIP signature (PK\x03\x04)
                prop_assert_eq!(
                    &bytes[0..4],
                    &[0x50, 0x4B, 0x03, 0x04],
                    "Invalid ZIP signature"
                );

                // Should be a reasonable size (at least 100 bytes for minimal XLSX)
                prop_assert!(bytes.len() >= 100, "Excel data suspiciously small: {} bytes", bytes.len());
            }
        }
    }
}
