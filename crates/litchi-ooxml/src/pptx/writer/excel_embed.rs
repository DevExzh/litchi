//! Excel workbook embedding for chart data.
//!
//! This module generates minimal XLSX files containing chart data for embedding
//! in PowerPoint presentations. The generated files are valid Excel workbooks
//! that PowerPoint can read to display chart data.

use crate::error::{OoxmlError, Result};
use crate::pptx::parts::chart::ChartData;
use litchi_core::xml::escape_xml;

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
    use litchi_opc::phys_pkg::PhysPkgWriter;

    crate::pptx::parts::chart::validate_chart_data(chart)?;

    let mut writer = PhysPkgWriter::new();

    // [Content_Types].xml
    let content_types = generate_content_types();
    write_part(
        &mut writer,
        "/[Content_Types].xml",
        content_types.as_bytes(),
    )?;

    // _rels/.rels
    let rels = generate_root_rels();
    write_part(&mut writer, "/_rels/.rels", rels.as_bytes())?;

    // xl/workbook.xml
    let workbook = generate_workbook_xml();
    write_part(&mut writer, "/xl/workbook.xml", workbook.as_bytes())?;

    // xl/_rels/workbook.xml.rels
    let workbook_rels = generate_workbook_rels();
    write_part(
        &mut writer,
        "/xl/_rels/workbook.xml.rels",
        workbook_rels.as_bytes(),
    )?;

    // xl/worksheets/sheet1.xml
    let sheet = generate_worksheet_xml(chart);
    write_part(&mut writer, "/xl/worksheets/sheet1.xml", sheet.as_bytes())?;

    // xl/styles.xml (minimal styles)
    let styles = generate_styles_xml();
    write_part(&mut writer, "/xl/styles.xml", styles.as_bytes())?;

    writer.finish().map_err(OoxmlError::from)
}

fn write_part(
    writer: &mut litchi_opc::phys_pkg::PhysPkgWriter,
    name: &'static str,
    data: &[u8],
) -> Result<()> {
    let name = litchi_opc::PackURI::new(name).map_err(OoxmlError::InvalidUri)?;
    writer.write(&name, data).map_err(OoxmlError::from)
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

    if matches!(
        chart.chart_type,
        crate::pptx::parts::chart::ChartType::Scatter
            | crate::pptx::parts::chart::ChartType::Bubble
    ) {
        return generate_xy_worksheet_xml(chart);
    }

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

fn generate_xy_worksheet_xml(chart: &ChartData) -> String {
    use crate::pptx::parts::chart::ChartType;
    use std::fmt::Write;

    let is_bubble = chart.chart_type == ChartType::Bubble;
    let columns_per_series = if is_bubble { 3 } else { 2 };
    let num_data_rows = chart
        .series
        .iter()
        .map(|series| {
            series
                .values
                .len()
                .max(series.x_values.len())
                .max(series.bubble_sizes.len())
                .max(series.categories.len())
        })
        .max()
        .unwrap_or(0);

    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
    xml.push_str(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#,
    );
    xml.push_str(r#"<row r="1">"#);
    for (series_index, series) in chart.series.iter().enumerate() {
        let first_column = series_index * columns_per_series;
        for (offset, suffix) in [(0, " X"), (1, "")] {
            let column = column_letter(first_column + offset);
            let _ = write!(
                xml,
                r#"<c r="{}1" t="inlineStr"><is><t>{}{}</t></is></c>"#,
                column,
                escape_xml(&series.name),
                suffix
            );
        }
        if is_bubble {
            let column = column_letter(first_column + 2);
            let _ = write!(
                xml,
                r#"<c r="{}1" t="inlineStr"><is><t>{} Size</t></is></c>"#,
                column,
                escape_xml(&series.name)
            );
        }
    }
    xml.push_str("</row>");

    for row_index in 0..num_data_rows {
        let row_number = row_index + 2;
        let _ = write!(xml, r#"<row r="{}">"#, row_number);
        for (series_index, series) in chart.series.iter().enumerate() {
            let first_column = series_index * columns_per_series;
            let x_value = series.x_values.get(row_index).copied().or_else(|| {
                if is_bubble {
                    None
                } else {
                    Some(
                        series
                            .categories
                            .get(row_index)
                            .and_then(|value| value.parse::<f64>().ok())
                            .unwrap_or(row_index as f64),
                    )
                }
            });
            for (offset, value) in [
                (0, x_value),
                (1, series.values.get(row_index).copied()),
                (
                    2,
                    is_bubble
                        .then(|| series.bubble_sizes.get(row_index).copied())
                        .flatten(),
                ),
            ] {
                if offset >= columns_per_series {
                    continue;
                }
                if let Some(value) = value {
                    let column = column_letter(first_column + offset);
                    let _ = write!(
                        xml,
                        r#"<c r="{}{}"><v>{}</v></c>"#,
                        column, row_number, value
                    );
                }
            }
        }
        xml.push_str("</row>");
    }

    xml.push_str("</sheetData></worksheet>");
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::parts::chart::{ChartSeries, ChartType};

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
        use litchi_opc::{OpcPackage, PackURI};

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
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let workbook = package
            .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap();
        assert!(
            workbook
                .blob()
                .windows(7)
                .any(|window| window == b"<sheet ")
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
                .is_ok()
        );
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
        assert!(!xml.contains(r#"t="s""#)); // No dangling shared-string reference
    }

    #[test]
    fn bubble_worksheet_uses_x_y_and_size_columns() {
        let chart = ChartData::new(ChartType::Bubble, 0, 0, 100, 100).add_series(
            ChartSeries::new("Reach")
                .with_x_values(vec![1.0, 2.0])
                .with_values(vec![10.0, 20.0])
                .with_bubble_sizes(vec![4.0, 8.0]),
        );

        let xml = generate_worksheet_xml(&chart);
        assert!(xml.contains(r#"r="A2"><v>1</v>"#));
        assert!(xml.contains(r#"r="B2"><v>10</v>"#));
        assert!(xml.contains(r#"r="C2"><v>4</v>"#));
        assert!(xml.contains("Reach Size"));
    }

    #[test]
    fn scatter_worksheet_allocates_a_pair_of_columns_per_series() {
        let chart = ChartData::new(ChartType::Scatter, 0, 0, 100, 100)
            .add_series(
                ChartSeries::new("First")
                    .with_x_values(vec![1.0])
                    .with_values(vec![2.0]),
            )
            .add_series(
                ChartSeries::new("Second")
                    .with_x_values(vec![3.0])
                    .with_values(vec![4.0]),
            );

        let xml = generate_worksheet_xml(&chart);
        assert!(xml.contains(r#"r="A2"><v>1</v>"#));
        assert!(xml.contains(r#"r="B2"><v>2</v>"#));
        assert!(xml.contains(r#"r="C2"><v>3</v>"#));
        assert!(xml.contains(r#"r="D2"><v>4</v>"#));
    }

    #[test]
    fn embedded_workbook_rejects_invalid_chart_dimensions() {
        let chart = ChartData::new(ChartType::Bubble, 0, 0, 100, 100).add_series(
            ChartSeries::new("Bad")
                .with_x_values(vec![1.0])
                .with_values(vec![2.0, 3.0])
                .with_bubble_sizes(vec![4.0]),
        );
        assert!(generate_chart_excel_data(&chart).is_err());
    }

    /// **Feature: charts-smartart-integration, Property 5: Excel data is valid OPC**
    /// **Validates: Requirements 1.3**
    ///
    /// For any `ChartData`, the generated bytes form an OPC package containing
    /// the required workbook and worksheet parts.
    #[cfg(test)]
    mod property_tests {
        use super::*;
        use litchi_opc::{OpcPackage, PackURI};
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
            fn prop_excel_data_is_valid_opc(chart in chart_data_strategy()) {
                // Generate Excel data
                let result = generate_chart_excel_data(&chart);

                // Should always succeed for valid input
                prop_assert!(result.is_ok(), "Excel generation failed: {:?}", result.err());

                let bytes = result.unwrap();

                let package = OpcPackage::from_bytes(&bytes)
                    .map_err(|error| TestCaseError::fail(format!("invalid OPC package: {error}")))?;
                for name in ["/xl/workbook.xml", "/xl/worksheets/sheet1.xml", "/xl/styles.xml"] {
                    let name = PackURI::new(name).unwrap();
                    prop_assert!(package.get_part(&name).is_ok(), "missing required part {name}");
                }
            }
        }
    }
}
