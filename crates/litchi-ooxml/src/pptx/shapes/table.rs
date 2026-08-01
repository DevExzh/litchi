use crate::error::{OoxmlError, Result};
use crate::pptx::shapes::textframe::{extract_drawingml_text, scan_drawingml_element_ranges};
/// Table shape implementation for PowerPoint presentations.
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const DRAWINGML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWINGML_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

/// Match a DrawingML element, tolerating fragments whose `a` prefix is
/// inherited from a surrounding document root (shape XML is stored without
/// namespace declarations).
fn matches_drawingml_name(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    local_name: &[u8],
    fragment_prefix: &Option<Option<Vec<u8>>>,
) -> bool {
    if name.local_name().as_ref() != local_name {
        return false;
    }
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
        },
        ResolveResult::Unknown(prefix) => {
            prefix.as_slice() == b"a"
                || fragment_prefix
                    .as_ref()
                    .and_then(|prefix| prefix.as_deref())
                    == Some(prefix.as_slice())
        },
        ResolveResult::Unbound => fragment_prefix == &Some(None),
    }
}

/// Table style and conditional-formatting switches (`a:tblPr`).
///
/// The booleans mirror the `a:tblPr` attributes of ECMA-376 Part 1,
/// 21.1.2.1.7 and select which conditional parts of the referenced table
/// style apply. All default to `false` when the attribute is omitted.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct TableProperties {
    /// `firstRow` — apply the first (header) row part style.
    pub first_row: bool,
    /// `firstCol` — apply the first column part style.
    pub first_col: bool,
    /// `lastRow` — apply the last (totals) row part style.
    pub last_row: bool,
    /// `lastCol` — apply the last column part style.
    pub last_col: bool,
    /// `bandRow` — apply horizontal (row) banding.
    pub band_row: bool,
    /// `bandCol` — apply vertical (column) banding.
    pub band_col: bool,
    /// The referenced table style GUID (`a:tableStyleId` text), as stored.
    pub style_id: Option<String>,
}

/// A table in a PowerPoint presentation.
///
/// Tables in PowerPoint are DrawingML tables (a:tbl) contained within
/// graphic frames. They contain rows, which contain cells.
///
/// # Examples
///
/// ```rust,ignore
/// if let Some(table) = graphic_frame.table() {
///     println!("Table: {}x{}", table.row_count()?, table.column_count()?);
///     
///     for (row_idx, row) in table.rows()?.iter().enumerate() {
///         for (col_idx, cell) in row.cells()?.iter().enumerate() {
///             println!("Cell[{},{}]: {}", row_idx, col_idx, cell.text()?);
///         }
///     }
/// }
/// ```
#[derive(Debug, Clone)]
pub struct Table {
    /// Raw XML bytes for the table
    xml_bytes: Vec<u8>,
}

impl Table {
    /// Create a new Table from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Extract table XML from graphic frame XML.
    ///
    /// GraphicFrames contain the table within their structure, so we need
    /// to extract just the table portion.
    pub fn from_graphic_frame_xml(xml_bytes: &[u8]) -> Result<Self> {
        let mut table = None;
        scan_drawingml_element_ranges(xml_bytes, b"tbl", |start, length| {
            if table.is_none() {
                let start = start as usize;
                table = Some(Table::new(
                    xml_bytes[start..start + length as usize].to_vec(),
                ));
            }
            Ok(())
        })?;
        table
            .ok_or_else(|| OoxmlError::PartNotFound("Table not found in graphic frame".to_string()))
    }

    /// Get the number of rows in the table.
    pub fn row_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_drawingml_element_ranges(&self.xml_bytes, b"tr", |_, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get the number of columns in the table.
    ///
    /// Returns the number of cells in the first row, or 0 if the table is empty.
    pub fn column_count(&self) -> Result<usize> {
        let rows = self.rows()?;
        if let Some(first_row) = rows.first() {
            first_row.cell_count()
        } else {
            Ok(0)
        }
    }

    /// Get all rows in the table.
    pub fn rows(&self) -> Result<Vec<TableRow>> {
        let mut rows = Vec::new();
        scan_drawingml_element_ranges(&self.xml_bytes, b"tr", |start, length| {
            let start = start as usize;
            rows.push(TableRow::new(
                self.xml_bytes[start..start + length as usize].to_vec(),
            ));
            Ok(())
        })?;
        Ok(rows)
    }

    /// Get a specific cell by row and column index.
    ///
    /// Indices are zero-based. Returns None if the indices are out of bounds.
    pub fn cell(&self, row_idx: usize, col_idx: usize) -> Result<Option<TableCell>> {
        let rows = self.rows()?;
        if let Some(row) = rows.get(row_idx) {
            let cells = row.cells()?;
            Ok(cells.get(col_idx).cloned())
        } else {
            Ok(None)
        }
    }

    /// Get the table style switches and referenced table style (`a:tblPr`).
    ///
    /// Returns `None` when the underlying XML is not a DrawingML table or
    /// declares no `a:tblPr` element.
    pub fn properties(&self) -> Result<Option<TableProperties>> {
        parse_table_properties(&self.xml_bytes)
    }
}

/// A row in a PowerPoint table.
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Raw XML bytes for this row
    xml_bytes: Vec<u8>,
}

impl TableRow {
    /// Create a new TableRow from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_drawingml_element_ranges(&self.xml_bytes, b"tc", |_, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get all cells in this row.
    pub fn cells(&self) -> Result<Vec<TableCell>> {
        let mut cells = Vec::new();
        scan_drawingml_element_ranges(&self.xml_bytes, b"tc", |start, length| {
            let start = start as usize;
            cells.push(TableCell::new(
                self.xml_bytes[start..start + length as usize].to_vec(),
            ));
            Ok(())
        })?;
        Ok(cells)
    }
}

/// A cell in a PowerPoint table.
#[derive(Debug, Clone)]
pub struct TableCell {
    /// Raw XML bytes for this cell
    xml_bytes: Vec<u8>,
}

impl TableCell {
    /// Create a new TableCell from XML bytes.
    pub fn new(xml_bytes: Vec<u8>) -> Self {
        Self { xml_bytes }
    }

    /// Extract all text from this cell.
    pub fn text(&self) -> Result<String> {
        Ok(extract_drawingml_text(&self.xml_bytes, Some(' '))?
            .trim()
            .to_string())
    }
}

/// Parse the `a:tblPr` properties of a DrawingML table (`a:tbl` root).
///
/// Returns `None` when the bytes do not form a DrawingML table or the table
/// carries no `a:tblPr` element.
fn parse_table_properties(xml_bytes: &[u8]) -> Result<Option<TableProperties>> {
    let mut reader = NsReader::from_reader(xml_bytes);
    let mut fragment_prefix: Option<Option<Vec<u8>>> = None;
    let mut depth = 0usize;
    let mut properties: Option<TableProperties> = None;
    let mut properties_depth: Option<usize> = None;
    let mut style_id_depth: Option<usize> = None;
    let mut style_id_text = String::new();

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("table XML is too deep".to_string())
                })?;
                if depth == 1
                    && !matches_drawingml_name(&namespace, element.name(), b"tbl", &fragment_prefix)
                {
                    // Not a table fragment (for example a bare a:tc); no
                    // properties are reported.
                    return Ok(None);
                }
                if depth == 2
                    && matches_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                    )
                {
                    if properties.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "table has multiple tblPr elements".to_string(),
                        ));
                    }
                    properties = Some(TableProperties {
                        first_row: on_off_attribute(&element, b"firstRow", decoder)?,
                        first_col: on_off_attribute(&element, b"firstCol", decoder)?,
                        last_row: on_off_attribute(&element, b"lastRow", decoder)?,
                        last_col: on_off_attribute(&element, b"lastCol", decoder)?,
                        band_row: on_off_attribute(&element, b"bandRow", decoder)?,
                        band_col: on_off_attribute(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                    properties_depth = Some(depth);
                } else if properties_depth == Some(depth - 1)
                    && matches_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                    )
                    && style_id_depth.replace(depth).is_some()
                {
                    return Err(OoxmlError::InvalidFormat(
                        "table has multiple tableStyleId elements".to_string(),
                    ));
                }
            },
            Event::Empty(element) => {
                if fragment_prefix.is_none() && !matches!(namespace, ResolveResult::Bound(_)) {
                    fragment_prefix = Some(
                        element
                            .name()
                            .prefix()
                            .map(|prefix| prefix.into_inner().to_vec()),
                    );
                }
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    OoxmlError::InvalidFormat("table XML is too deep".to_string())
                })?;
                if child_depth == 1
                    && !matches_drawingml_name(&namespace, element.name(), b"tbl", &fragment_prefix)
                {
                    return Ok(None);
                }
                if child_depth == 2
                    && matches_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                    )
                {
                    if properties.is_some() {
                        return Err(OoxmlError::InvalidFormat(
                            "table has multiple tblPr elements".to_string(),
                        ));
                    }
                    properties = Some(TableProperties {
                        first_row: on_off_attribute(&element, b"firstRow", decoder)?,
                        first_col: on_off_attribute(&element, b"firstCol", decoder)?,
                        last_row: on_off_attribute(&element, b"lastRow", decoder)?,
                        last_col: on_off_attribute(&element, b"lastCol", decoder)?,
                        band_row: on_off_attribute(&element, b"bandRow", decoder)?,
                        band_col: on_off_attribute(&element, b"bandCol", decoder)?,
                        style_id: None,
                    });
                }
            },
            Event::Text(text) if style_id_depth.is_some() => {
                let decoded = text
                    .decode()
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?;
                style_id_text.push_str(&decoded);
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(OoxmlError::InvalidFormat(
                        "invalid table XML nesting".to_string(),
                    ));
                }
                if style_id_depth == Some(depth)
                    && matches_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tableStyleId",
                        &fragment_prefix,
                    )
                {
                    let style_id = style_id_text.trim();
                    if style_id.is_empty() {
                        return Err(OoxmlError::InvalidFormat(
                            "tableStyleId must not be empty".to_string(),
                        ));
                    }
                    if let Some(properties) = properties.as_mut() {
                        properties.style_id = Some(style_id.to_string());
                    }
                    style_id_depth = None;
                    style_id_text.clear();
                }
                if properties_depth == Some(depth)
                    && matches_drawingml_name(
                        &namespace,
                        element.name(),
                        b"tblPr",
                        &fragment_prefix,
                    )
                {
                    properties_depth = None;
                }
                depth -= 1;
            },
            Event::Eof => {
                if depth != 0 || properties_depth.is_some() || style_id_depth.is_some() {
                    return Err(OoxmlError::InvalidFormat(
                        "unterminated table XML".to_string(),
                    ));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(properties)
}

/// Read one `ST_OnOff` attribute (`true/false/1/0/on/off`, default `false`).
fn on_off_attribute(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    match unqualified_attribute_value(element, name, decoder)? {
        None => Ok(false),
        Some(value) => match value.as_str() {
            "true" | "1" | "on" => Ok(true),
            "false" | "0" | "off" => Ok(false),
            _ => Err(OoxmlError::InvalidFormat(format!(
                "invalid table property value '{value}' for '{}'",
                String::from_utf8_lossy(name)
            ))),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn table_cell_text_decodes_runs_and_separates_paragraphs() {
        let cell = TableCell::new(
            br#"<d:tc xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><d:txBody><d:p><d:r><d:t>A &amp; </d:t></d:r><d:r><d:t><![CDATA[B < C]]></d:t></d:r></d:p><d:p><d:r><d:t>D</d:t></d:r></d:p></d:txBody></d:tc>"#
                .to_vec(),
        );
        assert_eq!(cell.text().unwrap(), "A & B < C D");
    }

    #[test]
    fn table_ranges_preserve_aliases_and_ignore_foreign_lookalikes() {
        let xml = br#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:false="urn:not-drawingml">
            <false:tbl><false:tr><false:tc/></false:tr></false:tbl>
            <d:tbl data-id="kept">
                <false:tr><false:tc/></false:tr>
                <d:tr h="20"><false:tc/><d:tc><d:txBody><d:p><d:r><d:t><![CDATA[A < B]]></d:t></d:r></d:p></d:txBody></d:tc></d:tr>
            </d:tbl>
        </p:graphicFrame>"#;
        let table = Table::from_graphic_frame_xml(xml).unwrap();
        assert_eq!(table.row_count().unwrap(), 1);
        let rows = table.rows().unwrap();
        assert_eq!(rows.len(), 1);
        assert_eq!(rows[0].cell_count().unwrap(), 1);
        let cells = rows[0].cells().unwrap();
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].text().unwrap(), "A < B");
    }

    #[test]
    fn table_ranges_reject_unterminated_selected_elements() {
        let xml = br#"<a:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tbl><a:tr/>"#;
        assert!(Table::from_graphic_frame_xml(xml).is_err());
    }

    #[test]
    fn table_properties_parse_switches_and_style_id() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblPr firstRow="1" lastCol="on" bandRow="true"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid/><a:tr/></a:tbl>"#;
        let properties = Table::new(xml.to_vec()).properties().unwrap().unwrap();
        assert!(properties.first_row);
        assert!(!properties.first_col);
        assert!(!properties.last_row);
        assert!(properties.last_col);
        assert!(properties.band_row);
        assert!(!properties.band_col);
        assert_eq!(
            properties.style_id.as_deref(),
            Some("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        );
    }

    #[test]
    fn table_properties_default_to_false_without_attributes() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblPr/><a:tblGrid/></a:tbl>"#;
        let properties = Table::new(xml.to_vec()).properties().unwrap().unwrap();
        assert_eq!(properties, TableProperties::default());
    }

    #[test]
    fn table_properties_are_absent_without_tblpr_or_table_root() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblGrid/></a:tbl>"#;
        assert!(Table::new(xml.to_vec()).properties().unwrap().is_none());
        let cell = br#"<a:tc xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>"#;
        assert!(Table::new(cell.to_vec()).properties().unwrap().is_none());
    }

    #[test]
    fn table_properties_reject_invalid_values_and_duplicates() {
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblPr firstRow="maybe"/></a:tbl>"#;
        assert!(Table::new(xml.to_vec()).properties().is_err());
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblPr/><a:tblPr/></a:tbl>"#;
        assert!(Table::new(xml.to_vec()).properties().is_err());
        let xml = br#"<a:tbl xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:tblPr><a:tableStyleId> </a:tableStyleId></a:tblPr></a:tbl>"#;
        assert!(Table::new(xml.to_vec()).properties().is_err());
    }

    #[test]
    fn table_properties_tolerate_undeclared_fragment_prefixes() {
        // Shape XML is stored without namespace declarations; the `a` prefix
        // is inherited from the surrounding slide root.
        let xml = br#"<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr><a:tblGrid/></a:tbl>"#;
        let properties = Table::new(xml.to_vec()).properties().unwrap().unwrap();
        assert!(properties.first_row);
        assert!(properties.band_row);
        assert_eq!(
            properties.style_id.as_deref(),
            Some("{5940675A-B579-460E-94D1-54222C63F5DA}")
        );
    }
}
