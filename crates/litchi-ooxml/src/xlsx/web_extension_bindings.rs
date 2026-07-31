//! Worksheet-side bindings for Office web extensions (MS-XLSX 2.2.4.12).

use std::collections::HashSet;
use std::ops::Range;

use litchi_core::sheet::Result;
use litchi_core::xml::escape_xml;
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use litchi_ooxml_common::xml::decode_xml_reference;

pub const WEB_EXTENSIONS_EXTENSION_URI: &str = "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}";
pub const X15_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
pub const XM_NAMESPACE: &str = "http://schemas.microsoft.com/office/excel/2006/main";

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_BINDINGS: usize = 65_536;
const MAX_STRING_BYTES: usize = 32_767;
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;

/// A worksheet range connected to an MS-OWEXML binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetWebExtensionBinding {
    application_reference: String,
    range_formula: String,
}

impl WorksheetWebExtensionBinding {
    pub fn new(
        application_reference: impl Into<String>,
        range_formula: impl Into<String>,
    ) -> Result<Self> {
        let value = Self {
            application_reference: application_reference.into(),
            range_formula: range_formula.into(),
        };
        value.validate()?;
        Ok(value)
    }

    pub fn application_reference(&self) -> &str {
        &self.application_reference
    }

    pub fn range_formula(&self) -> &str {
        &self.range_formula
    }

    fn validate(&self) -> Result<()> {
        bounded_nonempty(&self.application_reference, "web-extension appRef")?;
        bounded_nonempty(&self.range_formula, "web-extension range formula")?;
        validate_sheet_qualified_range(&self.range_formula)
    }
}

/// Parse the x15 web-extension binding collection embedded in worksheet `extLst`.
pub fn parse_worksheet_web_extension_bindings(
    worksheet_xml: &[u8],
) -> Result<Vec<WorksheetWebExtensionBinding>> {
    if worksheet_xml.len() > MAX_XML_BYTES {
        return Err("worksheet XML exceeds the web-extension parser limit".into());
    }
    let mut reader = NsReader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut bindings = Vec::new();
    let mut app_refs = HashSet::new();
    let mut depth = 0usize;
    let mut in_extension_list = false;
    let mut in_extension = false;
    let mut saw_extension = false;
    let mut in_collection = false;
    let mut saw_collection = false;
    let mut binding: Option<(String, String, bool)> = None;
    let mut in_formula = false;

    loop {
        let event = reader.read_event()?;
        match event {
            Event::DocType(_) => return Err("DTD is forbidden in worksheet XML".into()),
            Event::Start(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if local == b"extLst" && is_sml(&namespace) && depth == 1 {
                    in_extension_list = true;
                } else if local == b"ext" && is_sml(&namespace) && in_extension_list && depth == 2 {
                    let uri = attribute(&element, b"uri", reader.decoder())?;
                    if uri.as_deref() == Some(WEB_EXTENSIONS_EXTENSION_URI) {
                        if saw_extension {
                            return Err("duplicate worksheet webExtensions extension".into());
                        }
                        in_extension = true;
                        saw_extension = true;
                    }
                } else if in_extension
                    && namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtensions"
                    && depth == 3
                {
                    if saw_collection {
                        return Err("duplicate x15:webExtensions collection".into());
                    }
                    reject_attributes(&element)?;
                    in_collection = true;
                    saw_collection = true;
                } else if in_collection
                    && namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtension"
                    && depth == 4
                {
                    if binding.is_some() {
                        return Err("nested x15:webExtension".into());
                    }
                    let app_ref = attribute(&element, b"appRef", reader.decoder())?
                        .ok_or("x15:webExtension requires appRef")?;
                    reject_other_attributes(&element, b"appRef")?;
                    binding = Some((app_ref, String::new(), false));
                } else if binding.is_some()
                    && namespace == XM_NAMESPACE.as_bytes()
                    && local == b"f"
                    && depth == 5
                {
                    if in_formula || binding.as_ref().is_some_and(|binding| binding.2) {
                        return Err("x15:webExtension requires exactly one xm:f".into());
                    }
                    reject_attributes(&element)?;
                    in_formula = true;
                } else if binding.is_some() {
                    return Err("unexpected element in x15:webExtension".into());
                }
                depth += 1;
            },
            Event::Empty(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if in_collection
                    && namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtension"
                    && depth == 4
                {
                    return Err("x15:webExtension requires exactly one xm:f".into());
                }
                if binding.is_some() {
                    return Err("unexpected empty element in x15:webExtension".into());
                }
            },
            Event::Text(text) => {
                if in_formula {
                    let (_, formula, _) = binding.as_mut().expect("formula is inside binding");
                    formula.push_str(&text.decode()?);
                    if formula.len() > MAX_STRING_BYTES {
                        return Err("web-extension range formula is too long".into());
                    }
                } else if binding.is_some() && !text.decode()?.trim().is_empty() {
                    return Err("non-whitespace text outside xm:f".into());
                }
            },
            Event::CData(text) => {
                if in_formula {
                    let (_, formula, _) = binding.as_mut().expect("formula is inside binding");
                    formula.push_str(&text.decode()?);
                    if formula.len() > MAX_STRING_BYTES {
                        return Err("web-extension range formula is too long".into());
                    }
                } else if binding.is_some() && !text.decode()?.trim().is_empty() {
                    return Err("CDATA outside xm:f".into());
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or("unbalanced worksheet XML element")?;
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if namespace == XM_NAMESPACE.as_bytes() && local == b"f" && depth == 5 {
                    in_formula = false;
                    binding.as_mut().expect("formula is inside binding").2 = true;
                } else if namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtension"
                    && depth == 4
                {
                    let (app_ref, formula, has_formula) =
                        binding.take().ok_or("orphan x15:webExtension end")?;
                    if !has_formula {
                        return Err("x15:webExtension requires exactly one xm:f".into());
                    }
                    let value = WorksheetWebExtensionBinding::new(app_ref, formula)?;
                    if !app_refs.insert(value.application_reference.clone()) {
                        return Err("duplicate worksheet web-extension appRef".into());
                    }
                    bindings.push(value);
                    if bindings.len() > MAX_BINDINGS {
                        return Err("too many worksheet web-extension bindings".into());
                    }
                } else if namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtensions"
                    && depth == 3
                {
                    if bindings.is_empty() {
                        return Err("x15:webExtensions requires at least one binding".into());
                    }
                    in_collection = false;
                } else if is_sml(&namespace) && local == b"ext" && depth == 2 && in_extension {
                    if !saw_collection {
                        return Err("webExtensions extension omits x15:webExtensions".into());
                    }
                    in_extension = false;
                } else if is_sml(&namespace) && local == b"extLst" && depth == 1 {
                    in_extension_list = false;
                }
            },
            Event::GeneralRef(reference) if in_formula => {
                let (_, formula, _) = binding.as_mut().expect("formula is inside binding");
                formula.push_str(&decode_xml_reference(&reference)?);
                if formula.len() > MAX_STRING_BYTES {
                    return Err("web-extension range formula is too long".into());
                }
            },
            Event::GeneralRef(_) if binding.is_some() => {
                return Err("XML reference outside xm:f".into());
            },
            Event::GeneralRef(_) => {},
            Event::Eof => break,
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {},
        }
    }
    if depth != 0
        || binding.is_some()
        || in_formula
        || in_collection
        || in_extension
        || in_extension_list
    {
        return Err("unterminated worksheet web-extension binding XML".into());
    }
    Ok(bindings)
}

/// Serialize the complete SpreadsheetML `ext` element for worksheet bindings.
pub fn write_worksheet_web_extension_bindings(
    bindings: &[WorksheetWebExtensionBinding],
) -> Result<Vec<u8>> {
    write_bindings_extension(bindings, SML)
}

/// Replace or remove the binding extension without rebuilding unrelated XML.
pub fn replace_worksheet_web_extension_bindings(
    worksheet_xml: &[u8],
    bindings: &[WorksheetWebExtensionBinding],
) -> Result<Vec<u8>> {
    // Validate the selected vocabulary before doing byte-span mutation.
    let _ = parse_worksheet_web_extension_bindings(worksheet_xml)?;
    let scan = scan_extension_spans(worksheet_xml)?;
    if bindings.is_empty() {
        return match scan.extension {
            Some(range) => apply_edit(worksheet_xml, range, &[]),
            None => Ok(worksheet_xml.to_vec()),
        };
    }
    let extension = write_bindings_extension(bindings, &scan.spreadsheet_namespace)?;
    if let Some(range) = scan.extension {
        return apply_edit(worksheet_xml, range, &extension);
    }
    if let Some(position) = scan.ext_list_close {
        return apply_edit(worksheet_xml, position..position, &extension);
    }
    let position = scan
        .worksheet_close
        .ok_or("worksheet document has no closing worksheet element")?;
    let mut wrapper = Vec::with_capacity(extension.len() + 32);
    wrapper.extend_from_slice(b"<extLst>");
    wrapper.extend_from_slice(&extension);
    wrapper.extend_from_slice(b"</extLst>");
    apply_edit(worksheet_xml, position..position, &wrapper)
}

/// Require every worksheet `appRef` to resolve to one package binding.
pub fn validate_worksheet_web_extension_apprefs(
    worksheet_bindings: &[WorksheetWebExtensionBinding],
    package_bindings: &[crate::web_extensions::WebExtensionBinding],
) -> Result<()> {
    let mut package_refs = HashSet::with_capacity(package_bindings.len());
    for binding in package_bindings {
        if !package_refs.insert(binding.application_reference.as_str()) {
            return Err("duplicate MS-OWEXML binding appref".into());
        }
    }
    for binding in worksheet_bindings {
        if !package_refs.contains(binding.application_reference()) {
            return Err(format!(
                "worksheet web-extension appRef '{}' has no MS-OWEXML binding",
                binding.application_reference()
            )
            .into());
        }
    }
    Ok(())
}

fn write_bindings_extension(
    bindings: &[WorksheetWebExtensionBinding],
    spreadsheet_namespace: &str,
) -> Result<Vec<u8>> {
    if bindings.is_empty() {
        return Err("x15:webExtensions requires at least one binding".into());
    }
    if bindings.len() > MAX_BINDINGS {
        return Err("too many worksheet web-extension bindings".into());
    }
    let mut seen = HashSet::with_capacity(bindings.len());
    let mut xml = format!(
        r#"<ext xmlns="{spreadsheet_namespace}" uri="{WEB_EXTENSIONS_EXTENSION_URI}"><x15:webExtensions xmlns:x15="{X15_NAMESPACE}" xmlns:xm="{XM_NAMESPACE}">"#
    );
    for binding in bindings {
        binding.validate()?;
        if !seen.insert(&binding.application_reference) {
            return Err("duplicate worksheet web-extension appRef".into());
        }
        xml.push_str(r#"<x15:webExtension appRef=""#);
        xml.push_str(&escape_xml(&binding.application_reference));
        xml.push_str(r#""><xm:f>"#);
        xml.push_str(&escape_xml(&binding.range_formula));
        xml.push_str("</xm:f></x15:webExtension>");
    }
    xml.push_str("</x15:webExtensions></ext>");
    Ok(xml.into_bytes())
}

#[derive(Debug)]
struct ExtensionScan {
    spreadsheet_namespace: String,
    extension: Option<Range<usize>>,
    ext_list_close: Option<usize>,
    worksheet_close: Option<usize>,
}

fn scan_extension_spans(xml: &[u8]) -> Result<ExtensionScan> {
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut root_namespace = None;
    let mut matching_start = None;
    let mut extension = None;
    let mut ext_list_close = None;
    let mut worksheet_close = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| "worksheet XML offset exceeds usize")?;
        let event = reader.read_event()?;
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if depth == 0 && local == b"worksheet" && is_sml(&namespace) {
                    root_namespace = Some(String::from_utf8(namespace.clone())?);
                }
                if local == b"ext"
                    && is_sml(&namespace)
                    && depth == 2
                    && attribute(&element, b"uri", reader.decoder())?.as_deref()
                        == Some(WEB_EXTENSIONS_EXTENSION_URI)
                    && (matching_start.replace((start, depth)).is_some() || extension.is_some())
                {
                    return Err("duplicate worksheet webExtensions extension".into());
                }
                depth += 1;
            },
            Event::Empty(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if local == b"ext"
                    && is_sml(&namespace)
                    && depth == 2
                    && attribute(&element, b"uri", reader.decoder())?.as_deref()
                        == Some(WEB_EXTENSIONS_EXTENSION_URI)
                {
                    if extension.is_some() || matching_start.is_some() {
                        return Err("duplicate worksheet webExtensions extension".into());
                    }
                    extension = Some(start..usize::try_from(reader.buffer_position()).unwrap());
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or("unbalanced worksheet XML element")?;
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if matching_start.is_some_and(|(_, target_depth)| target_depth == depth)
                    && local == b"ext"
                    && is_sml(&namespace)
                {
                    let (begin, _) = matching_start.take().expect("matching extension checked");
                    extension = Some(
                        begin
                            ..usize::try_from(reader.buffer_position())
                                .map_err(|_| "worksheet XML offset exceeds usize")?,
                    );
                }
                if local == b"extLst" && is_sml(&namespace) && depth == 1 {
                    ext_list_close = Some(start);
                }
                if local == b"worksheet" && is_sml(&namespace) && depth == 0 {
                    worksheet_close = Some(start);
                }
            },
            Event::DocType(_) => return Err("DTD is forbidden in worksheet XML".into()),
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(ExtensionScan {
        spreadsheet_namespace: root_namespace.ok_or("worksheet root element is missing")?,
        extension,
        ext_list_close,
        worksheet_close,
    })
}

fn apply_edit(xml: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > xml.len() {
        return Err("invalid worksheet XML edit span".into());
    }
    let mut output = Vec::with_capacity(xml.len() - (range.end - range.start) + replacement.len());
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}

fn validate_sheet_qualified_range(value: &str) -> Result<()> {
    if value.trim() != value {
        return Err("web-extension formula cannot contain surrounding whitespace".into());
    }
    let bang = find_unquoted_bang(value)?;
    let (sheet, range) = value.split_at(bang);
    validate_sheet_name(sheet)?;
    validate_a1_range(&range[1..])
}

fn find_unquoted_bang(value: &str) -> Result<usize> {
    let mut quoted = false;
    let bytes = value.as_bytes();
    let mut index = 0;
    let mut found = None;
    while index < bytes.len() {
        match bytes[index] {
            b'\'' => {
                if quoted && bytes.get(index + 1) == Some(&b'\'') {
                    index += 1;
                } else {
                    quoted = !quoted;
                }
            },
            b'!' if !quoted && found.is_some() => {
                return Err("web-extension formula must contain one sheet qualifier".into());
            },
            b'!' if !quoted => found = Some(index),
            _ => {},
        }
        index += 1;
    }
    if quoted {
        return Err("unterminated quoted worksheet name".into());
    }
    found.ok_or_else(|| "web-extension formula requires a sheet qualifier".into())
}

fn validate_sheet_name(value: &str) -> Result<()> {
    if value.is_empty() {
        return Err("empty worksheet name in web-extension formula".into());
    }
    if value.starts_with('\'') {
        if !value.ends_with('\'') || value.len() < 3 {
            return Err("invalid quoted worksheet name".into());
        }
    } else if value
        .bytes()
        .any(|byte| !byte.is_ascii_alphanumeric() && byte != b'_' && byte != b'.')
    {
        return Err("worksheet name must be quoted in web-extension formula".into());
    }
    Ok(())
}

fn validate_a1_range(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    validate_a1_cell(parts.next().unwrap_or_default())?;
    if let Some(last) = parts.next() {
        validate_a1_cell(last)?;
    }
    if parts.next().is_some() {
        return Err("web-extension formula contains more than one range operator".into());
    }
    Ok(())
}

fn validate_a1_cell(value: &str) -> Result<()> {
    let value = value.strip_prefix('$').unwrap_or(value);
    let column_end = value
        .bytes()
        .position(|byte| !byte.is_ascii_alphabetic())
        .unwrap_or(value.len());
    if column_end == 0 || column_end > 3 {
        return Err("invalid column in web-extension range".into());
    }
    let column = value[..column_end].bytes().fold(0u32, |number, byte| {
        number * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1)
    });
    let row = value[column_end..]
        .strip_prefix('$')
        .unwrap_or(&value[column_end..]);
    if column > 16_384
        || row.is_empty()
        || !row.bytes().all(|byte| byte.is_ascii_digit())
        || row
            .parse::<u32>()
            .ok()
            .filter(|row| (1..=1_048_576).contains(row))
            .is_none()
    {
        return Err("invalid cell in web-extension range".into());
    }
    Ok(())
}

fn bounded_nonempty(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_STRING_BYTES {
        return Err(format!("{field} must contain 1..={MAX_STRING_BYTES} bytes").into());
    }
    if value.chars().any(|character| character.is_control()) {
        return Err(format!("{field} contains a control character").into());
    }
    Ok(())
}

fn resolved_name(
    reader: &NsReader<&[u8]>,
    name: quick_xml::name::QName<'_>,
) -> Result<(Vec<u8>, Vec<u8>)> {
    let namespace = match reader.resolver().resolve_element(name).0 {
        ResolveResult::Bound(namespace) => namespace.as_ref().to_vec(),
        ResolveResult::Unbound => Vec::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(format!("unknown XML namespace prefix {:?}", prefix).into());
        },
    };
    Ok((namespace, name.local_name().as_ref().to_vec()))
}

fn is_sml(namespace: &[u8]) -> bool {
    namespace == SML.as_bytes() || namespace == STRICT_SML.as_bytes()
}

fn attribute(
    element: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute?;
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if value.is_some() {
                return Err("duplicate XML attribute".into());
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn reject_attributes(element: &quick_xml::events::BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err("unexpected XML attribute".into());
        }
    }
    Ok(())
}

fn reject_other_attributes(
    element: &quick_xml::events::BytesStart<'_>,
    permitted: &[u8],
) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute?;
        if !is_namespace_declaration(attribute.key.as_ref())
            && (attribute.key.prefix().is_some()
                || attribute.key.local_name().as_ref() != permitted)
        {
            return Err("unexpected x15:webExtension attribute".into());
        }
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_fixture_and_roundtrips_canonical_xml() {
        let xml =
            include_bytes!("../../../../test-data/ooxml/web_extensions/worksheet_bindings.xml");
        let parsed = parse_worksheet_web_extension_bindings(xml).unwrap();
        assert_eq!(parsed.len(), 2);
        assert_eq!(parsed[0].range_formula(), "Sheet1!$A$1:$B$4");
        assert_eq!(parsed[1].range_formula(), "'Sales 2026'!C3");
        let encoded = write_worksheet_web_extension_bindings(&parsed).unwrap();
        let reparsed = parse_worksheet_web_extension_bindings(&[
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst>"#
                .as_slice(),
            encoded.as_slice(),
            b"</extLst></worksheet>",
        ]
        .concat())
        .unwrap();
        assert_eq!(reparsed, parsed);
    }

    #[test]
    fn validates_range_grammar_and_unique_apprefs() {
        for invalid in [
            "A1",
            "Sheet1!A0",
            "Sheet1!XFE1",
            "Sheet1!A1:B2:C3",
            "A!B!C1",
        ] {
            assert!(
                WorksheetWebExtensionBinding::new("binding", invalid).is_err(),
                "{invalid}"
            );
        }
        let binding = WorksheetWebExtensionBinding::new("same", "Sheet1!A1").unwrap();
        assert!(write_worksheet_web_extension_bindings(&[binding.clone(), binding]).is_err());
    }

    #[test]
    fn rejects_wrong_grammar_and_malformed_xml() {
        let prefix = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst><ext uri="{F7C9EE02-42E1-4005-9D12-6889AFFD525C}"><x15:webExtensions xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#;
        let suffix = b"</x15:webExtensions></ext></extLst></worksheet>";
        for body in [
            br#"<x15:webExtension appRef="a"/>"#.as_slice(),
            br#"<x15:webExtension appRef="a"><xm:f>A1</xm:f></x15:webExtension>"#,
            br#"<x15:webExtension appRef="a"><xm:f>Sheet1!A1</xm:f><xm:f>Sheet1!A2</xm:f></x15:webExtension>"#,
        ] {
            assert!(
                parse_worksheet_web_extension_bindings(&[prefix.as_slice(), body, suffix].concat())
                    .is_err()
            );
        }
        let duplicate = [
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst>"#
                .as_slice(),
            br#"<ext uri="{F7C9EE02-42E1-4005-9D12-6889AFFD525C}"><x15:webExtensions xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x15:webExtension appRef="a"><xm:f>Sheet1!A1</xm:f></x15:webExtension></x15:webExtensions></ext>"#,
            br#"<ext uri="{F7C9EE02-42E1-4005-9D12-6889AFFD525C}"><x15:webExtensions xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x15:webExtension appRef="b"><xm:f>Sheet1!A2</xm:f></x15:webExtension></x15:webExtensions></ext>"#,
            b"</extLst></worksheet>",
        ]
        .concat();
        assert!(parse_worksheet_web_extension_bindings(&duplicate).is_err());
    }

    #[test]
    fn replaces_inserts_and_removes_without_touching_unrelated_xml() {
        let source = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData><extLst><ext uri="opaque"><foreign/></ext></extLst></worksheet>"#;
        let binding = WorksheetWebExtensionBinding::new("binding", "Sheet1!A1").unwrap();
        let inserted =
            replace_worksheet_web_extension_bindings(source, std::slice::from_ref(&binding))
                .unwrap();
        assert!(
            inserted
                .windows(b"<row r=\"1\"/>".len())
                .any(|w| w == b"<row r=\"1\"/>")
        );
        assert!(
            inserted
                .windows(b"uri=\"opaque\"".len())
                .any(|w| w == b"uri=\"opaque\"")
        );
        assert_eq!(
            parse_worksheet_web_extension_bindings(&inserted).unwrap(),
            [binding]
        );
        let removed = replace_worksheet_web_extension_bindings(&inserted, &[]).unwrap();
        assert!(
            parse_worksheet_web_extension_bindings(&removed)
                .unwrap()
                .is_empty()
        );
        assert!(
            removed
                .windows(b"uri=\"opaque\"".len())
                .any(|w| w == b"uri=\"opaque\"")
        );
    }

    #[test]
    fn validates_package_appref_links() {
        let worksheet = [WorksheetWebExtensionBinding::new("binding", "Sheet1!A1").unwrap()];
        let package = [crate::web_extensions::WebExtensionBinding {
            id: "id".into(),
            binding_type: "table".into(),
            application_reference: "binding".into(),
            extension_list: None,
        }];
        validate_worksheet_web_extension_apprefs(&worksheet, &package).unwrap();
        assert!(validate_worksheet_web_extension_apprefs(&worksheet, &[]).is_err());
    }
}
