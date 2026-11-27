//! The members of a minimal workbook holding an authored chart's data.

use std::fmt::Write as _;

use super::{Cell, Spec, column_name};
use crate::Result;

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#;
const ROOT_RELATIONSHIPS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#;
const WORKBOOK: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
const WORKBOOK_RELATIONSHIPS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>"#;

fn worksheet(cells: &[(usize, usize, Cell<'_>)]) -> String {
    let mut xml = String::with_capacity(4096);
    xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>"#);
    let mut row = None;
    for (index, column, cell) in cells {
        if row != Some(*index) {
            if row.is_some() {
                xml.push_str("</row>");
            }
            let _ = write!(xml, r#"<row r="{}">"#, index + 1);
            row = Some(*index);
        }
        let reference = format!("{}{}", column_name(*column), index + 1);
        let _ = match cell {
            Cell::Text(text) => write!(
                xml,
                r#"<c r="{reference}" t="inlineStr"><is><t>{}</t></is></c>"#,
                escape(text)
            ),
            Cell::Number(value) => write!(xml, r#"<c r="{reference}"><v>{value}</v></c>"#),
        };
    }
    if row.is_some() {
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    xml
}

impl Spec {
    /// The members of a minimal workbook whose `Sheet1` holds this spec's
    /// [`Self::sheet_layout`] at `A1`, as `(part name, bytes)` in a fixed
    /// order: the data an embedded chart workbook carries. The host zips
    /// them into the package it embeds; the same spec always produces the
    /// same members.
    ///
    /// ```
    /// use litchi_drawingml::chart::authoring::{Kind, Series, Spec, Variant};
    ///
    /// let spec = Spec::new(Variant::new(Kind::Column))
    ///     .with_categories(["Q1"])
    ///     .with_series(Series::new("East", vec![1.0]));
    /// let members = spec.workbook_members()?;
    /// assert_eq!(members[0].0, "/[Content_Types].xml");
    /// let sheet = String::from_utf8(members[5].1.clone()).expect("UTF-8");
    /// assert!(sheet.contains(r#"<c r="B2"><v>1</v></c>"#));
    /// # Ok::<(), litchi_drawingml::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::validate`].
    pub fn workbook_members(&self) -> Result<Vec<(&'static str, Vec<u8>)>> {
        self.validate()?;
        Ok(vec![
            ("/[Content_Types].xml", CONTENT_TYPES.as_bytes().to_vec()),
            ("/_rels/.rels", ROOT_RELATIONSHIPS.as_bytes().to_vec()),
            ("/xl/workbook.xml", WORKBOOK.as_bytes().to_vec()),
            (
                "/xl/_rels/workbook.xml.rels",
                WORKBOOK_RELATIONSHIPS.as_bytes().to_vec(),
            ),
            ("/xl/styles.xml", STYLES.as_bytes().to_vec()),
            (
                "/xl/worksheets/sheet1.xml",
                worksheet(&self.sheet_layout()).into_bytes(),
            ),
        ])
    }
}
