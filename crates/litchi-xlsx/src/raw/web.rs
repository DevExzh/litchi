//! Physical `SpreadsheetML` codec for worksheet Office web-extension bindings.

use std::ops::Range;

use litchi_ooxml_common::xml::decode_xml_reference;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

use crate::error::{Result, allocation, invalid};
use crate::web::{Binding, Bindings, MAX_BINDINGS, MAX_STRING_BYTES};

/// `SpreadsheetML` extension URI for worksheet web-extension bindings.
pub const EXTENSION_URI: &str = "{F7C9EE02-42E1-4005-9D12-6889AFFD525C}";
/// Namespace of the worksheet web-extension collection.
pub const X15_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
/// Namespace of the range formula element.
pub const XM_NAMESPACE: &str = "http://schemas.microsoft.com/office/excel/2006/main";

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 128;

/// Read the checked worksheet binding collection embedded in `extLst`.
pub fn read(worksheet_xml: &[u8]) -> Result<Bindings> {
    if worksheet_xml.len() > MAX_XML_BYTES {
        return Err(invalid(
            "worksheet XML exceeds the web-extension parser limit",
        ));
    }
    let mut reader = NsReader::from_reader(worksheet_xml);
    reader.config_mut().trim_text(false);
    let mut values = Vec::new();
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;
    let mut in_extension_list = false;
    let mut in_extension = false;
    let mut saw_extension = false;
    let mut in_collection = false;
    let mut saw_collection = false;
    let mut binding: Option<(String, String, bool)> = None;
    let mut in_formula = false;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?;
        match event {
            Event::DocType(_) => return Err(invalid("DTD is forbidden in worksheet XML")),
            Event::Start(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if depth == 0 {
                    if saw_root || closed_root || local != b"worksheet" || !is_sml(&namespace) {
                        return Err(invalid(
                            "worksheet XML must have one SpreadsheetML worksheet root",
                        ));
                    }
                    saw_root = true;
                } else if local == b"extLst" && is_sml(&namespace) && depth == 1 {
                    in_extension_list = true;
                } else if local == b"ext" && is_sml(&namespace) && in_extension_list && depth == 2 {
                    let uri = attribute(&element, b"uri", reader.decoder())?;
                    if uri.as_deref() == Some(EXTENSION_URI) {
                        if saw_extension {
                            return Err(invalid("duplicate worksheet webExtensions extension"));
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
                        return Err(invalid("duplicate x15:webExtensions collection"));
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
                        return Err(invalid("nested x15:webExtension"));
                    }
                    let app_ref = attribute(&element, b"appRef", reader.decoder())?
                        .ok_or_else(|| invalid("x15:webExtension requires appRef"))?;
                    reject_other_attributes(&element, b"appRef")?;
                    binding = Some((app_ref, String::new(), false));
                } else if binding.is_some()
                    && namespace == XM_NAMESPACE.as_bytes()
                    && local == b"f"
                    && depth == 5
                {
                    if in_formula || binding.as_ref().is_some_and(|value| value.2) {
                        return Err(invalid("x15:webExtension requires exactly one xm:f"));
                    }
                    reject_attributes(&element)?;
                    in_formula = true;
                } else if binding.is_some() {
                    return Err(invalid("unexpected element in x15:webExtension"));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid(format!(
                        "worksheet XML exceeds the {MAX_DEPTH} element depth limit"
                    )));
                }
            },
            Event::Empty(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if depth == 0 {
                    if saw_root || closed_root || local != b"worksheet" || !is_sml(&namespace) {
                        return Err(invalid(
                            "worksheet XML must have one SpreadsheetML worksheet root",
                        ));
                    }
                    saw_root = true;
                    closed_root = true;
                } else if local == b"ext"
                    && is_sml(&namespace)
                    && in_extension_list
                    && depth == 2
                    && attribute(&element, b"uri", reader.decoder())?.as_deref()
                        == Some(EXTENSION_URI)
                {
                    return Err(invalid("webExtensions extension omits x15:webExtensions"));
                } else if in_collection
                    && namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtension"
                    && depth == 4
                {
                    return Err(invalid("x15:webExtension requires exactly one xm:f"));
                }
                if binding.is_some() {
                    return Err(invalid("unexpected empty element in x15:webExtension"));
                }
            },
            Event::Text(text) => {
                let text = text.decode().map_err(|error| invalid(error.to_string()))?;
                if in_formula {
                    push_formula(&mut binding, &text)?;
                } else if binding.is_some() && !text.trim().is_empty() {
                    return Err(invalid("non-whitespace text outside xm:f"));
                }
            },
            Event::CData(text) => {
                let text = text.decode().map_err(|error| invalid(error.to_string()))?;
                if in_formula {
                    push_formula(&mut binding, &text)?;
                } else if binding.is_some() && !text.trim().is_empty() {
                    return Err(invalid("CDATA outside xm:f"));
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced worksheet XML element"))?;
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if namespace == XM_NAMESPACE.as_bytes() && local == b"f" && depth == 5 {
                    in_formula = false;
                    binding
                        .as_mut()
                        .ok_or_else(|| invalid("orphan xm:f end element"))?
                        .2 = true;
                } else if namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtension"
                    && depth == 4
                {
                    let (app_ref, formula, has_formula) = binding
                        .take()
                        .ok_or_else(|| invalid("orphan x15:webExtension end"))?;
                    if !has_formula {
                        return Err(invalid("x15:webExtension requires exactly one xm:f"));
                    }
                    if values.len() >= MAX_BINDINGS {
                        return Err(invalid("too many worksheet web-extension bindings"));
                    }
                    values
                        .try_reserve(1)
                        .map_err(|source| allocation("worksheet web bindings", source))?;
                    values.push(Binding::new(app_ref, formula)?);
                } else if namespace == X15_NAMESPACE.as_bytes()
                    && local == b"webExtensions"
                    && depth == 3
                {
                    if values.is_empty() {
                        return Err(invalid("x15:webExtensions requires at least one binding"));
                    }
                    in_collection = false;
                } else if is_sml(&namespace) && local == b"ext" && depth == 2 && in_extension {
                    if !saw_collection {
                        return Err(invalid("webExtensions extension omits x15:webExtensions"));
                    }
                    in_extension = false;
                } else if is_sml(&namespace) && local == b"extLst" && depth == 1 {
                    in_extension_list = false;
                } else if is_sml(&namespace) && local == b"worksheet" && depth == 0 {
                    closed_root = true;
                }
            },
            Event::GeneralRef(reference) if in_formula => {
                let decoded = decode_xml_reference(&reference)?;
                push_formula(&mut binding, &decoded)?;
            },
            Event::GeneralRef(_) if binding.is_some() => {
                return Err(invalid("XML reference outside xm:f"));
            },
            Event::GeneralRef(_) => {},
            Event::Eof => break,
            Event::Decl(_) | Event::PI(_) | Event::Comment(_) => {},
        }
    }
    if !saw_root || !closed_root || depth != 0 {
        return Err(invalid(
            "worksheet XML has a missing or unterminated SpreadsheetML root",
        ));
    }
    if binding.is_some() || in_formula || in_collection || in_extension || in_extension_list {
        return Err(invalid("unterminated worksheet web-extension binding XML"));
    }
    Bindings::try_from(values)
}

/// Write the complete transitional `SpreadsheetML` `ext` element.
pub fn write(bindings: &Bindings) -> Result<Vec<u8>> {
    write_for_namespace(bindings, SML)
}

/// Replace, insert, or remove the binding extension without rebuilding
/// unrelated worksheet XML.
pub fn replace(worksheet_xml: &[u8], bindings: &Bindings) -> Result<Vec<u8>> {
    // Validate the selected vocabulary before doing a byte-span mutation.
    let _ = read(worksheet_xml)?;
    let scan = scan_extension_spans(worksheet_xml)?;
    if bindings.is_empty() {
        return match scan.extension {
            Some(range) => apply_edit(worksheet_xml, range, &[]),
            None => copy_bytes(worksheet_xml),
        };
    }
    let extension = write_for_namespace(bindings, &scan.spreadsheet_namespace)?;
    if let Some(range) = scan.extension {
        return apply_edit(worksheet_xml, range, &extension);
    }
    if let Some(position) = scan.ext_list_close {
        return apply_edit(worksheet_xml, position..position, &extension);
    }
    let position = scan
        .worksheet_close
        .ok_or_else(|| invalid("worksheet document has no closing worksheet element"))?;
    let wrapper_len = extension
        .len()
        .checked_add(b"<extLst></extLst>".len())
        .ok_or_else(|| invalid("worksheet web-extension wrapper size overflow"))?;
    if wrapper_len > MAX_XML_BYTES {
        return Err(invalid("worksheet web-extension wrapper exceeds XML limit"));
    }
    let mut wrapper = Vec::new();
    wrapper
        .try_reserve_exact(wrapper_len)
        .map_err(|source| allocation("worksheet web-extension wrapper", source))?;
    wrapper.extend_from_slice(b"<extLst>");
    wrapper.extend_from_slice(&extension);
    wrapper.extend_from_slice(b"</extLst>");
    apply_edit(worksheet_xml, position..position, &wrapper)
}

fn push_formula(binding: &mut Option<(String, String, bool)>, value: &str) -> Result<()> {
    let (_, formula, _) = binding
        .as_mut()
        .ok_or_else(|| invalid("xm:f appeared outside x15:webExtension"))?;
    let length = formula
        .len()
        .checked_add(value.len())
        .ok_or_else(|| invalid("web-extension range formula size overflow"))?;
    if length > MAX_STRING_BYTES {
        return Err(invalid("web-extension range formula is too long"));
    }
    formula
        .try_reserve(value.len())
        .map_err(|source| allocation("web-extension range formula", source))?;
    formula.push_str(value);
    Ok(())
}

fn write_for_namespace(bindings: &Bindings, spreadsheet_namespace: &str) -> Result<Vec<u8>> {
    bindings.validate_all()?;
    if bindings.is_empty() {
        return Err(invalid("x15:webExtensions requires at least one binding"));
    }
    let mut length = b"<ext xmlns=\"\" uri=\"\"><x15:webExtensions xmlns:x15=\"\" xmlns:xm=\"\"></x15:webExtensions></ext>".len();
    length = checked_xml_len(length, spreadsheet_namespace.len())?;
    length = checked_xml_len(length, EXTENSION_URI.len())?;
    length = checked_xml_len(length, X15_NAMESPACE.len())?;
    length = checked_xml_len(length, XM_NAMESPACE.len())?;
    for binding in bindings {
        length = checked_xml_len(
            length,
            b"<x15:webExtension appRef=\"\"><xm:f></xm:f></x15:webExtension>".len(),
        )?;
        length = checked_xml_len(length, escaped_len(binding.app_ref())?)?;
        length = checked_xml_len(length, escaped_len(binding.formula())?)?;
    }

    let mut xml = String::new();
    xml.try_reserve_exact(length)
        .map_err(|source| allocation("worksheet web XML", source))?;
    xml.push_str("<ext xmlns=\"");
    xml.push_str(spreadsheet_namespace);
    xml.push_str("\" uri=\"");
    xml.push_str(EXTENSION_URI);
    xml.push_str("\"><x15:webExtensions xmlns:x15=\"");
    xml.push_str(X15_NAMESPACE);
    xml.push_str("\" xmlns:xm=\"");
    xml.push_str(XM_NAMESPACE);
    xml.push_str("\">");
    for binding in bindings {
        xml.push_str("<x15:webExtension appRef=\"");
        push_escaped(&mut xml, binding.app_ref());
        xml.push_str("\"><xm:f>");
        push_escaped(&mut xml, binding.formula());
        xml.push_str("</xm:f></x15:webExtension>");
    }
    xml.push_str("</x15:webExtensions></ext>");
    if xml.len() != length {
        return Err(invalid(
            "worksheet web-extension XML size calculation mismatch",
        ));
    }
    Ok(xml.into_bytes())
}

fn escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        checked_xml_len(
            length,
            match character {
                '&' => 5,
                '<' | '>' => 4,
                '\'' | '"' => 6,
                character => character.len_utf8(),
            },
        )
    })
}

fn push_escaped(output: &mut String, value: &str) {
    for character in value.chars() {
        output.push_str(match character {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '\'' => "&apos;",
            '"' => "&quot;",
            _ => {
                output.push(character);
                continue;
            },
        });
    }
}

fn checked_xml_len(current: usize, additional: usize) -> Result<usize> {
    let length = current
        .checked_add(additional)
        .ok_or_else(|| invalid("worksheet web-extension XML size overflow"))?;
    if length > MAX_XML_BYTES {
        return Err(invalid(format!(
            "worksheet web-extension XML exceeds the {MAX_XML_BYTES} byte limit"
        )));
    }
    Ok(length)
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
            .map_err(|_source| invalid("worksheet XML offset exceeds usize"))?;
        let event = reader
            .read_event()
            .map_err(|error| invalid(error.to_string()))?;
        match event {
            Event::Start(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if depth == 0 && local == b"worksheet" && is_sml(&namespace) {
                    root_namespace = Some(
                        String::from_utf8(namespace.clone())
                            .map_err(|error| invalid(error.to_string()))?,
                    );
                }
                if local == b"ext"
                    && is_sml(&namespace)
                    && depth == 2
                    && attribute(&element, b"uri", reader.decoder())?.as_deref()
                        == Some(EXTENSION_URI)
                {
                    if matching_start.is_some() || extension.is_some() {
                        return Err(invalid("duplicate worksheet webExtensions extension"));
                    }
                    matching_start = Some((start, depth));
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML depth overflow"))?;
            },
            Event::Empty(element) => {
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if local == b"ext"
                    && is_sml(&namespace)
                    && depth == 2
                    && attribute(&element, b"uri", reader.decoder())?.as_deref()
                        == Some(EXTENSION_URI)
                {
                    if extension.is_some() || matching_start.is_some() {
                        return Err(invalid("duplicate worksheet webExtensions extension"));
                    }
                    let end = usize::try_from(reader.buffer_position())
                        .map_err(|_source| invalid("worksheet XML offset exceeds usize"))?;
                    extension = Some(start..end);
                }
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced worksheet XML element"))?;
                let (namespace, local) = resolved_name(&reader, element.name())?;
                if matching_start.is_some_and(|(_, target_depth)| target_depth == depth)
                    && local == b"ext"
                    && is_sml(&namespace)
                {
                    let (begin, _) = matching_start.take().ok_or_else(|| {
                        invalid("worksheet webExtensions start element is missing")
                    })?;
                    extension = Some(
                        begin
                            ..usize::try_from(reader.buffer_position())
                                .map_err(|_source| invalid("worksheet XML offset exceeds usize"))?,
                    );
                }
                if local == b"extLst" && is_sml(&namespace) && depth == 1 {
                    ext_list_close = Some(start);
                }
                if local == b"worksheet" && is_sml(&namespace) && depth == 0 {
                    worksheet_close = Some(start);
                }
            },
            Event::DocType(_) => return Err(invalid("DTD is forbidden in worksheet XML")),
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(ExtensionScan {
        spreadsheet_namespace: root_namespace
            .ok_or_else(|| invalid("worksheet root element is missing"))?,
        extension,
        ext_list_close,
        worksheet_close,
    })
}

fn apply_edit(xml: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > xml.len() {
        return Err(invalid("invalid worksheet XML edit span"));
    }
    let removed = range.end - range.start;
    let length = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| invalid("worksheet XML edit size overflow"))?;
    if length > MAX_XML_BYTES {
        return Err(invalid(format!(
            "edited worksheet XML exceeds the {MAX_XML_BYTES} byte limit"
        )));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| allocation("edited worksheet XML", source))?;
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}

fn copy_bytes(value: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation("worksheet XML copy", source))?;
    output.extend_from_slice(value);
    Ok(output)
}

fn resolved_name(reader: &NsReader<&[u8]>, name: QName<'_>) -> Result<(Vec<u8>, Vec<u8>)> {
    let namespace = match reader.resolver().resolve_element(name).0 {
        ResolveResult::Bound(namespace) => namespace.as_ref().to_vec(),
        ResolveResult::Unbound => Vec::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!("unknown XML namespace prefix {prefix:?}")));
        },
    };
    Ok((namespace, name.local_name().as_ref().to_vec()))
}

fn is_sml(namespace: &[u8]) -> bool {
    namespace == SML.as_bytes() || namespace == STRICT_SML.as_bytes()
}

fn attribute(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if attribute.key.prefix().is_none() && attribute.key.local_name().as_ref() == name {
            if value.is_some() {
                return Err(invalid("duplicate XML attribute"));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                    .map_err(|error| invalid(error.to_string()))?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn reject_attributes(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if !is_namespace_declaration(attribute.key.as_ref()) {
            return Err(invalid("unexpected XML attribute"));
        }
    }
    Ok(())
}

fn reject_other_attributes(element: &BytesStart<'_>, permitted: &[u8]) -> Result<()> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(error.to_string()))?;
        if !is_namespace_declaration(attribute.key.as_ref())
            && (attribute.key.prefix().is_some()
                || attribute.key.local_name().as_ref() != permitted)
        {
            return Err(invalid("unexpected x15:webExtension attribute"));
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

    fn binding(app_ref: &str, formula: &str) -> Binding {
        Binding::new(app_ref, formula).unwrap()
    }

    #[test]
    fn parses_fixture_and_roundtrips_canonical_xml() {
        let xml =
            include_bytes!("../../../../test-data/ooxml/web_extensions/worksheet_bindings.xml");
        let parsed = read(xml).unwrap();
        assert_eq!(parsed.len(), 2);
        assert_eq!(
            parsed.get(0).map(Binding::formula),
            Some("Sheet1!$A$1:$B$4")
        );
        assert_eq!(parsed.get(1).map(Binding::formula), Some("'Sales 2026'!C3"));
        let encoded = write(&parsed).unwrap();
        let reparsed = read(
            &[
                br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst>"#
                    .as_slice(),
                encoded.as_slice(),
                b"</extLst></worksheet>",
            ]
            .concat(),
        )
        .unwrap();
        assert_eq!(reparsed, parsed);
    }

    #[test]
    fn rejects_wrong_grammar_duplicates_and_empty_writes() {
        let prefix = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst><ext uri="{F7C9EE02-42E1-4005-9D12-6889AFFD525C}"><x15:webExtensions xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#;
        let suffix = b"</x15:webExtensions></ext></extLst></worksheet>";
        for body in [
            br#"<x15:webExtension appRef="a"/>"#.as_slice(),
            br#"<x15:webExtension appRef="a"><xm:f>A1</xm:f></x15:webExtension>"#,
            br#"<x15:webExtension appRef="a"><xm:f>Sheet1!A1</xm:f><xm:f>Sheet1!A2</xm:f></x15:webExtension>"#,
        ] {
            assert!(read(&[prefix.as_slice(), body, suffix].concat()).is_err());
        }
        let duplicate = [
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><extLst><ext uri="{F7C9EE02-42E1-4005-9D12-6889AFFD525C}"><x15:webExtensions xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">"#.as_slice(),
            br#"<x15:webExtension appRef="a"><xm:f>Sheet1!A1</xm:f></x15:webExtension><x15:webExtension appRef="a"><xm:f>Sheet1!A2</xm:f></x15:webExtension>"#,
            b"</x15:webExtensions></ext></extLst></worksheet>",
        ]
        .concat();
        assert!(read(&duplicate).is_err());
        assert!(write(&Bindings::new()).is_err());
    }

    #[test]
    fn replaces_inserts_and_removes_without_touching_unrelated_xml() {
        let source = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"/></sheetData><extLst><ext uri="opaque"><foreign/></ext></extLst></worksheet>"#;
        let bindings = Bindings::try_from(vec![binding("binding", "Sheet1!A1")]).unwrap();
        let inserted = replace(source, &bindings).unwrap();
        assert!(
            inserted
                .windows(b"<row r=\"1\"/>".len())
                .any(|value| value == b"<row r=\"1\"/>")
        );
        assert!(
            inserted
                .windows(b"uri=\"opaque\"".len())
                .any(|value| value == b"uri=\"opaque\"")
        );
        assert_eq!(read(&inserted).unwrap(), bindings);

        let removed = replace(&inserted, &Bindings::new()).unwrap();
        assert!(read(&removed).unwrap().is_empty());
        assert!(
            removed
                .windows(b"uri=\"opaque\"".len())
                .any(|value| value == b"uri=\"opaque\"")
        );
        assert_eq!(replace(&removed, &Bindings::new()).unwrap(), removed);
    }

    #[test]
    fn retains_strict_spreadsheetml_namespace() {
        let source = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData/></worksheet>"#;
        let bindings = Bindings::try_from(vec![binding("strict", "Sheet1!A1")]).unwrap();
        let replaced = replace(source, &bindings).unwrap();
        let text = std::str::from_utf8(&replaced).unwrap();
        assert!(text.contains("xmlns=\"http://purl.oclc.org/ooxml/spreadsheetml/main\""));
        assert_eq!(read(&replaced).unwrap(), bindings);
    }
}
