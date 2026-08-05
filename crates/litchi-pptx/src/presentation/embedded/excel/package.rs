use super::model::{Chart, Kind};
use crate::Result;
use litchi_opc::PackURI;
use litchi_opc::phys_pkg::PhysPkgWriter;
use std::fmt::Write;

/// Generate a minimal, valid XLSX package containing one worksheet of chart
/// data. The chart itself is intentionally not coupled to a PPTX chart part.
pub fn generate(chart: &Chart) -> Result<Vec<u8>> {
    chart.validate()?;
    let mut writer = PhysPkgWriter::new();
    write_part(
        &mut writer,
        "/[Content_Types].xml",
        content_types().as_bytes(),
    )?;
    write_part(&mut writer, "/_rels/.rels", root_relationships().as_bytes())?;
    write_part(&mut writer, "/xl/workbook.xml", workbook().as_bytes())?;
    write_part(
        &mut writer,
        "/xl/_rels/workbook.xml.rels",
        workbook_relationships().as_bytes(),
    )?;
    write_part(&mut writer, "/xl/styles.xml", styles().as_bytes())?;
    let worksheet = worksheet(chart);
    write_part(
        &mut writer,
        "/xl/worksheets/sheet1.xml",
        worksheet.as_bytes(),
    )?;
    Ok(writer.finish()?)
}

fn write_part(writer: &mut PhysPkgWriter, name: &str, bytes: &[u8]) -> Result<()> {
    let name = PackURI::new(name).map_err(crate::Error::Uri)?;
    writer.write(&name, bytes)?;
    Ok(())
}

fn content_types() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#.to_string()
}

fn root_relationships() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#.to_string()
}

fn workbook() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_string()
}

fn workbook_relationships() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_string()
}

fn styles() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#.to_string()
}

fn worksheet(chart: &Chart) -> String {
    if matches!(chart.kind, Kind::Scatter | Kind::Bubble) {
        return xy_worksheet(chart);
    }
    let rows = chart
        .series
        .iter()
        .map(|series| series.values.len().max(series.categories.len()))
        .max()
        .unwrap_or(0);
    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#);
    for (index, series) in chart.series.iter().enumerate() {
        let _ = write!(
            xml,
            r#"<c r="{}1" t="inlineStr"><is><t>{}</t></is></c>"#,
            column(index + 1),
            escape(&series.name)
        );
    }
    xml.push_str("</row>");
    for row in 0..rows {
        let number = row + 2;
        let _ = write!(xml, r#"<row r="{number}">"#);
        if let Some(category) = chart
            .series
            .first()
            .and_then(|series| series.categories.get(row))
        {
            let _ = write!(
                xml,
                r#"<c r="A{number}" t="inlineStr"><is><t>{}</t></is></c>"#,
                escape(category)
            );
        }
        for (index, series) in chart.series.iter().enumerate() {
            if let Some(value) = series.values.get(row) {
                let _ = write!(
                    xml,
                    r#"<c r="{}{number}"><v>{value}</v></c>"#,
                    column(index + 1)
                );
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    xml
}

fn xy_worksheet(chart: &Chart) -> String {
    let bubble = chart.kind == Kind::Bubble;
    let stride = if bubble { 3 } else { 2 };
    let rows = chart
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
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#,
    );
    for (series_index, series) in chart.series.iter().enumerate() {
        let first = series_index * stride;
        for (offset, suffix) in if bubble {
            vec![(0, " X"), (1, ""), (2, " Size")]
        } else {
            vec![(0, " X"), (1, "")]
        } {
            let _ = write!(
                xml,
                r#"<c r="{}1" t="inlineStr"><is><t>{}{suffix}</t></is></c>"#,
                column(first + offset),
                escape(&series.name)
            );
        }
    }
    xml.push_str("</row>");
    for row in 0..rows {
        let number = row + 2;
        let _ = write!(xml, r#"<row r="{number}">"#);
        for (series_index, series) in chart.series.iter().enumerate() {
            let first = series_index * stride;
            let x_value = series.x_values.get(row).copied().or_else(|| {
                (!bubble).then(|| {
                    series
                        .categories
                        .get(row)
                        .and_then(|value| value.parse::<f64>().ok())
                        .unwrap_or(row as f64)
                })
            });
            if let Some(value) = x_value {
                let _ = write!(
                    xml,
                    r#"<c r="{}{number}"><v>{value}</v></c>"#,
                    column(first)
                );
            }
            if let Some(value) = series.values.get(row) {
                let _ = write!(
                    xml,
                    r#"<c r="{}{number}"><v>{value}</v></c>"#,
                    column(first + 1)
                );
            }
            if bubble && let Some(value) = series.bubble_sizes.get(row) {
                let _ = write!(
                    xml,
                    r#"<c r="{}{number}"><v>{value}</v></c>"#,
                    column(first + 2)
                );
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    xml
}

fn column(mut value: usize) -> String {
    let mut result = String::new();
    loop {
        result.insert(0, (b'A' + (value % 26) as u8) as char);
        if value < 26 {
            break;
        }
        value = value / 26 - 1;
    }
    result
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::OpcPackage;

    #[test]
    fn generates_a_valid_opaque_workbook() {
        let chart = Chart::new(Kind::Column, 0, 0, 100, 100).add_series(
            super::super::Series::new("Sales")
                .with_categories(vec!["Q1".into(), "Q2".into()])
                .with_values(vec![10.0, 20.0]),
        );
        let bytes = generate(&chart).unwrap();
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(
            package
                .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
                .is_ok()
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn rejects_invalid_bubble_shapes() {
        let chart = Chart::new(Kind::Bubble, 0, 0, 100, 100).add_series(
            super::super::Series::new("Bad")
                .with_x_values(vec![1.0])
                .with_values(vec![2.0, 3.0])
                .with_bubble_sizes(vec![4.0]),
        );
        assert!(generate(&chart).is_err());
    }

    #[test]
    fn keeps_categories_in_column_a_and_series_in_following_columns() {
        let chart = Chart::new(Kind::Column, 0, 0, 100, 100).add_series(
            super::super::Series::new("Sales")
                .with_categories(vec!["Q1".into(), "Q2".into()])
                .with_values(vec![10.0, 20.0]),
        );

        let xml = worksheet(&chart);
        assert!(xml.contains(r#"r="B1" t="inlineStr""#));
        assert!(xml.contains(r#"r="A2" t="inlineStr"#));
        assert!(xml.contains(r#"r="B2"><v>10</v>"#));
        assert!(!xml.contains(r#"r="A2"><v>10</v>"#));
    }

    #[test]
    fn scatter_charts_use_numeric_categories_when_x_values_are_absent() {
        let chart = Chart::new(Kind::Scatter, 0, 0, 100, 100).add_series(
            super::super::Series::new("Trend")
                .with_categories(vec!["2".into(), "not numeric".into()])
                .with_values(vec![10.0, 20.0]),
        );

        let xml = worksheet(&chart);
        assert!(xml.contains(r#"r="A2"><v>2</v>"#));
        assert!(xml.contains(r#"r="B2"><v>10</v>"#));
        assert!(xml.contains(r#"r="A3"><v>1</v>"#));
        assert!(xml.contains(r#"r="B3"><v>20</v>"#));
    }
}
