//! ODF spreadsheet, sheet, and cell protection metadata.

use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const LOEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
const OFFICE_EXT_NAMESPACE: &[u8] = b"http://openoffice.org/2009/office";

/// Stored protection key and digest-algorithm metadata.
///
/// These values are password verifiers, not encryption keys. Litchi preserves
/// them verbatim and does not attempt to recover or verify a password.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ProtectionKey {
    pub value: Option<String>,
    pub digest_algorithm: Option<String>,
    /// LibreOffice's secondary digest URI for legacy Excel-compatible hashes.
    pub secondary_digest_algorithm: Option<String>,
}

/// Protection metadata on the document-level `office:spreadsheet` element.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SpreadsheetProtection {
    /// `None` preserves an omitted attribute; `Some(false)` preserves an explicit false value.
    pub structure_protected: Option<bool>,
    pub key: ProtectionKey,
}

/// Granular edit permissions used by LibreOffice's table-protection extension.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SheetProtectionOptions {
    pub select_protected_cells: Option<bool>,
    pub select_unprotected_cells: Option<bool>,
    pub insert_columns: Option<bool>,
    pub insert_rows: Option<bool>,
    pub delete_columns: Option<bool>,
    pub delete_rows: Option<bool>,
    pub use_auto_filter: Option<bool>,
    pub use_pivot: Option<bool>,
}

impl SheetProtectionOptions {
    pub(crate) fn is_empty(&self) -> bool {
        self.select_protected_cells.is_none()
            && self.select_unprotected_cells.is_none()
            && self.insert_columns.is_none()
            && self.insert_rows.is_none()
            && self.delete_columns.is_none()
            && self.delete_rows.is_none()
            && self.use_auto_filter.is_none()
            && self.use_pivot.is_none()
    }
}

/// Protection metadata on a `table:table` sheet.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SheetProtection {
    pub protected: Option<bool>,
    pub key: ProtectionKey,
    pub options: SheetProtectionOptions,
}

pub(crate) fn parse_protection(xml: &str) -> Result<(SpreadsheetProtection, Vec<SheetProtection>)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut spreadsheet = SpreadsheetProtection::default();
    let mut spreadsheet_seen = false;
    let mut sheets = Vec::new();
    let mut current_sheet: Option<SheetProtection> = None;
    let mut element_depth = 0usize;
    let mut spreadsheet_depth = None;
    let mut current_sheet_depth = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let is_start = matches!(&event, Event::Start(_));
        let is_end = matches!(&event, Event::End(_));
        if is_start {
            element_depth += 1;
        }
        match event {
            Event::Start(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"spreadsheet" =>
            {
                if spreadsheet_seen {
                    return Err(Error::InvalidFormat(
                        "duplicate office:spreadsheet element".to_string(),
                    ));
                }
                spreadsheet = parse_spreadsheet_attributes(&reader, &element)?;
                spreadsheet_seen = true;
                spreadsheet_depth = Some(element_depth);
            },
            Event::Empty(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"spreadsheet" =>
            {
                if spreadsheet_seen {
                    return Err(Error::InvalidFormat(
                        "duplicate office:spreadsheet element".to_string(),
                    ));
                }
                spreadsheet = parse_spreadsheet_attributes(&reader, &element)?;
                spreadsheet_seen = true;
            },
            Event::Start(element)
                if is_namespace(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"table"
                    && spreadsheet_depth.is_some_and(|depth| element_depth == depth + 1) =>
            {
                current_sheet = Some(parse_sheet_attributes(&reader, &element)?);
                current_sheet_depth = Some(element_depth);
            },
            Event::Empty(element)
                if is_namespace(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"table"
                    && spreadsheet_depth == Some(element_depth)
                    && current_sheet.is_none() =>
            {
                sheets.push(parse_sheet_attributes(&reader, &element)?);
            },
            Event::Start(element) | Event::Empty(element)
                if current_sheet.is_some()
                    && current_sheet_depth.is_some_and(|depth| {
                        if is_start {
                            element_depth == depth + 1
                        } else {
                            element_depth == depth
                        }
                    })
                    && element.local_name().as_ref() == b"table-protection"
                    && is_protection_extension_namespace(&namespace) =>
            {
                let options = parse_sheet_options(&reader, &element)?;
                let sheet = current_sheet.as_mut().expect("checked sheet");
                if !sheet.options.is_empty() {
                    return Err(Error::InvalidFormat(
                        "duplicate sheet table-protection element".to_string(),
                    ));
                }
                sheet.options = options;
            },
            Event::End(element)
                if is_namespace(&namespace, TABLE_NAMESPACE)
                    && element.local_name().as_ref() == b"table"
                    && current_sheet_depth == Some(element_depth) =>
            {
                sheets.push(current_sheet.take().expect("checked sheet"));
                current_sheet_depth = None;
            },
            Event::End(element)
                if is_namespace(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"spreadsheet" =>
            {
                spreadsheet_depth = None;
            },
            Event::Eof => break,
            _ => {},
        }
        if is_end {
            element_depth = element_depth.saturating_sub(1);
        }
        buffer.clear();
    }
    if current_sheet.is_some() || current_sheet_depth.is_some() || spreadsheet_depth.is_some() {
        return Err(Error::InvalidFormat(
            "unterminated protected sheet".to_string(),
        ));
    }
    Ok((spreadsheet, sheets))
}

fn parse_spreadsheet_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SpreadsheetProtection> {
    Ok(SpreadsheetProtection {
        structure_protected: optional_bool_attribute(
            reader,
            element,
            TABLE_NAMESPACE,
            b"structure-protected",
        )?,
        key: parse_key(reader, element)?,
    })
}

fn parse_sheet_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SheetProtection> {
    Ok(SheetProtection {
        protected: optional_bool_attribute(reader, element, TABLE_NAMESPACE, b"protected")?,
        key: parse_key(reader, element)?,
        options: SheetProtectionOptions::default(),
    })
}

fn parse_key(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<ProtectionKey> {
    Ok(ProtectionKey {
        value: optional_attribute(reader, element, TABLE_NAMESPACE, b"protection-key")?,
        digest_algorithm: optional_attribute(
            reader,
            element,
            TABLE_NAMESPACE,
            b"protection-key-digest-algorithm",
        )?,
        secondary_digest_algorithm: optional_attribute(
            reader,
            element,
            LOEXT_NAMESPACE,
            b"protection-key-digest-algorithm-2",
        )?,
    })
}

fn parse_sheet_options(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<SheetProtectionOptions> {
    Ok(SheetProtectionOptions {
        select_protected_cells: optional_extension_bool(
            reader,
            element,
            b"select-protected-cells",
        )?,
        select_unprotected_cells: optional_extension_bool(
            reader,
            element,
            b"select-unprotected-cells",
        )?,
        insert_columns: optional_extension_bool(reader, element, b"insert-columns")?,
        insert_rows: optional_extension_bool(reader, element, b"insert-rows")?,
        delete_columns: optional_extension_bool(reader, element, b"delete-columns")?,
        delete_rows: optional_extension_bool(reader, element, b"delete-rows")?,
        use_auto_filter: optional_extension_bool(reader, element, b"use-autofilter")?,
        use_pivot: optional_extension_bool(reader, element, b"use-pivot")?,
    })
}

fn optional_extension_bool(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local_name: &[u8],
) -> Result<Option<bool>> {
    for namespace in [TABLE_NAMESPACE, LOEXT_NAMESPACE, OFFICE_EXT_NAMESPACE] {
        if let Some(value) = optional_bool_attribute(reader, element, namespace, local_name)? {
            return Ok(Some(value));
        }
    }
    Ok(None)
}

pub(crate) fn write_spreadsheet_attributes(out: &mut String, value: &SpreadsheetProtection) {
    write_bool_attribute(out, "table:structure-protected", value.structure_protected);
    write_key_attributes(out, &value.key);
}

pub(crate) fn write_sheet_attributes(out: &mut String, value: &SheetProtection) {
    write_bool_attribute(out, "table:protected", value.protected);
    write_key_attributes(out, &value.key);
}

pub(crate) fn write_sheet_options(out: &mut String, value: &SheetProtectionOptions) {
    if value.is_empty() {
        return;
    }
    out.push_str("<loext:table-protection");
    write_bool_attribute(
        out,
        "loext:select-protected-cells",
        value.select_protected_cells,
    );
    write_bool_attribute(
        out,
        "loext:select-unprotected-cells",
        value.select_unprotected_cells,
    );
    write_bool_attribute(out, "loext:insert-columns", value.insert_columns);
    write_bool_attribute(out, "loext:insert-rows", value.insert_rows);
    write_bool_attribute(out, "loext:delete-columns", value.delete_columns);
    write_bool_attribute(out, "loext:delete-rows", value.delete_rows);
    write_bool_attribute(out, "loext:use-autofilter", value.use_auto_filter);
    write_bool_attribute(out, "loext:use-pivot", value.use_pivot);
    out.push_str("/>");
}

pub(crate) fn has_extensions<'a>(
    spreadsheet: &SpreadsheetProtection,
    mut sheets: impl Iterator<Item = &'a SheetProtection>,
) -> bool {
    spreadsheet.key.secondary_digest_algorithm.is_some()
        || sheets.any(|sheet| {
            sheet.key.secondary_digest_algorithm.is_some() || !sheet.options.is_empty()
        })
}

fn write_key_attributes(out: &mut String, key: &ProtectionKey) {
    write_attribute(out, "table:protection-key", key.value.as_deref());
    write_attribute(
        out,
        "table:protection-key-digest-algorithm",
        key.digest_algorithm.as_deref(),
    );
    write_attribute(
        out,
        "loext:protection-key-digest-algorithm-2",
        key.secondary_digest_algorithm.as_deref(),
    );
}

fn write_bool_attribute(out: &mut String, name: &str, value: Option<bool>) {
    write_attribute(
        out,
        name,
        value.map(|value| if value { "true" } else { "false" }),
    );
}

fn write_attribute(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}

fn optional_bool_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<bool>> {
    optional_attribute(reader, element, namespace, local_name)?
        .map(|value| match value.as_str() {
            "true" | "1" => Ok(true),
            "false" | "0" => Ok(false),
            _ => Err(Error::InvalidFormat(format!(
                "invalid protection Boolean value '{value}'"
            ))),
        })
        .transpose()
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace_uri: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if is_namespace(&namespace, namespace_uri) && local.as_ref() == local_name {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")));
        }
    }
    Ok(None)
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn is_protection_extension_namespace(namespace: &ResolveResult<'_>) -> bool {
    [TABLE_NAMESPACE, LOEXT_NAMESPACE, OFFICE_EXT_NAMESPACE]
        .iter()
        .any(|expected| is_namespace(namespace, expected))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_and_writes_standard_and_libreoffice_protection() {
        let xml = r#"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:l="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" t:structure-protected="true" t:protection-key="abc&amp;=" t:protection-key-digest-algorithm="urn:sha256" l:protection-key-digest-algorithm-2="urn:sha1"><t:table t:name="Sheet1" t:protected="true" t:protection-key="sheet" t:protection-key-digest-algorithm="urn:sha1"><l:table-protection l:select-protected-cells="true" l:insert-rows="false" l:use-autofilter="true"/></t:table></o:spreadsheet>"#;
        let (spreadsheet, sheets) = parse_protection(xml).unwrap();
        assert_eq!(spreadsheet.structure_protected, Some(true));
        assert_eq!(spreadsheet.key.value.as_deref(), Some("abc&="));
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].protected, Some(true));
        assert_eq!(sheets[0].options.insert_rows, Some(false));

        let mut attributes = String::new();
        write_spreadsheet_attributes(&mut attributes, &spreadsheet);
        assert!(attributes.contains("table:structure-protected=\"true\""));
        assert!(attributes.contains("table:protection-key=\"abc&amp;=\""));
        let mut options = String::new();
        write_sheet_options(&mut options, &sheets[0].options);
        assert!(options.contains("loext:insert-rows=\"false\""));
    }

    #[test]
    fn rejects_invalid_boolean_and_ignores_nested_tables() {
        let invalid = r#"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" t:structure-protected="yes"/>"#;
        assert!(parse_protection(invalid).is_err());
        let nested = r#"<o:spreadsheet xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><t:table><t:table/></t:table></o:spreadsheet>"#;
        assert_eq!(parse_protection(nested).unwrap().1.len(), 1);
    }

    #[test]
    fn ignores_dde_cache_table_protection() {
        let xml = r#"<o:spreadsheet
          xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
          xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
          <t:dde-links><t:dde-link>
            <o:dde-source o:dde-application="app" o:dde-topic="topic" o:dde-item="item"/>
            <t:table t:name="Cache" t:protected="true"/>
          </t:dde-link></t:dde-links>
          <t:table t:name="Visible" t:protected="false"/>
        </o:spreadsheet>"#;

        let (_, sheets) = parse_protection(xml).unwrap();
        assert_eq!(sheets.len(), 1);
        assert_eq!(sheets[0].protected, Some(false));
    }
}
