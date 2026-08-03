//! Immutable XLSX worksheet phonetic-properties read model.

use crate::error::{Error, Result};
use crate::raw::namespace::is_spreadsheetml_name;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_EVENTS: usize = 1_000_000;

/// East Asian character conversion used for worksheet phonetic text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetPhoneticType {
    HalfWidthKatakana,
    FullWidthKatakana,
    Hiragana,
    NoConversion,
}

impl WorksheetPhoneticType {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "halfwidthKatakana" => Ok(Self::HalfWidthKatakana),
            "fullwidthKatakana" => Ok(Self::FullWidthKatakana),
            "Hiragana" => Ok(Self::Hiragana),
            "noConversion" => Ok(Self::NoConversion),
            _ => Err(invalid(format!(
                "invalid worksheet phonetic type '{value}'"
            ))),
        }
    }
}

/// Alignment of phonetic text relative to its base text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetPhoneticAlignment {
    NoControl,
    Left,
    Center,
    Distributed,
}

impl WorksheetPhoneticAlignment {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "noControl" => Ok(Self::NoControl),
            "left" => Ok(Self::Left),
            "center" => Ok(Self::Center),
            "distributed" => Ok(Self::Distributed),
            _ => Err(invalid(format!(
                "invalid worksheet phonetic alignment '{value}'"
            ))),
        }
    }
}

/// Effective default formatting from a worksheet's direct `phoneticPr` child.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WorksheetPhoneticProperties {
    font_id: u32,
    phonetic_type: WorksheetPhoneticType,
    alignment: WorksheetPhoneticAlignment,
}

impl WorksheetPhoneticProperties {
    /// Zero-based font index. Out-of-range indices fall back to the Normal-style font.
    pub fn font_id(&self) -> u32 {
        self.font_id
    }
    pub fn phonetic_type(&self) -> WorksheetPhoneticType {
        self.phonetic_type
    }
    pub fn alignment(&self) -> WorksheetPhoneticAlignment {
        self.alignment
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Outside,
    Worksheet,
    PhoneticProperties,
}

struct Parser {
    stack: Vec<Context>,
    properties: Option<WorksheetPhoneticProperties>,
    seen_properties: bool,
}

/// Parse the worksheet's exact `worksheet/phoneticPr` child path.
pub fn parse_worksheet_phonetic_properties(
    xml: &[u8],
) -> Result<Option<WorksheetPhoneticProperties>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML is too large"));
    }
    let limits = MceLimits {
        max_input_bytes: MAX_XML_BYTES,
        max_output_bytes: MAX_XML_BYTES,
        max_depth: MAX_DEPTH,
        ..MceLimits::default()
    };
    let processed = process_markup_compatibility(xml, &MceCapabilities::default(), &limits)?;
    if processed.xml.len() > MAX_XML_BYTES {
        return Err(invalid("processed worksheet XML is too large"));
    }
    let mut reader = NsReader::from_reader(processed.xml.as_ref());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut parser = Parser {
        stack: Vec::new(),
        properties: None,
        seen_properties: false,
    };
    let mut root_seen = false;
    let mut root_closed = false;
    let mut declaration_seen = false;
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("worksheet XML event count overflow"))?;
        if events > MAX_EVENTS {
            return Err(invalid("worksheet XML exceeds event limit"));
        }
        let decoder = reader.decoder();
        let event = reader.read_event().map_err(xml_error)?.into_owned();
        reject_unsafe_event(&event)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if parser.stack.is_empty() && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                if parser.stack.len() >= MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                let was_root = parser.stack.is_empty();
                parser.start(&namespace, &element, decoder, &resolver)?;
                if was_root {
                    root_seen = true;
                }
            },
            Event::Empty(element) => parser.empty(&namespace, &element, decoder, &resolver)?,
            Event::End(element) => {
                parser.end(element.local_name().as_ref())?;
                if parser.stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::Text(text)
                if parser.parent() == Context::PhoneticProperties
                    && !text.decode().map_err(xml_error)?.trim().is_empty() =>
            {
                return Err(invalid("unexpected text in worksheet phoneticPr"));
            },
            Event::Text(text) => {
                if (!root_seen || root_closed) && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet XML text is outside root"));
                }
                if parser.parent() == Context::Worksheet
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(invalid("worksheet cannot contain direct text"));
                }
            },
            Event::CData(_) if parser.parent() == Context::PhoneticProperties => {
                return Err(invalid("unexpected CDATA in worksheet phoneticPr"));
            },
            Event::CData(_) if !root_seen || root_closed => {
                return Err(invalid("worksheet XML CDATA is outside root"));
            },
            Event::CData(_) if parser.parent() == Context::Worksheet => {
                return Err(invalid("worksheet cannot contain direct CDATA"));
            },
            Event::GeneralRef(_) => {
                if !root_seen || root_closed {
                    return Err(invalid("worksheet XML entity is outside root"));
                }
                if parser.parent() == Context::Worksheet
                    || parser.parent() == Context::PhoneticProperties
                {
                    return Err(invalid("worksheet phoneticPr cannot contain entity text"));
                }
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !root_seen || !root_closed || !parser.stack.is_empty() {
        return Err(invalid("unterminated worksheet phoneticPr XML"));
    }
    Ok(parser.properties)
}

impl Parser {
    fn parent(&self) -> Context {
        self.stack.last().copied().unwrap_or(Context::Outside)
    }

    fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() && (!core || local.as_ref() != b"worksheet") {
            return Err(invalid("phoneticPr parser requires a worksheet root"));
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Outside, true, b"worksheet") => self.stack.push(Context::Worksheet),
            (Context::Worksheet, true, b"phoneticPr") => {
                self.add_properties(element, decoder, resolver)?;
                self.stack.push(Context::PhoneticProperties);
            },
            (Context::PhoneticProperties, _, _) => {
                return Err(invalid("worksheet phoneticPr is a leaf element"));
            },
            _ => self.stack.push(Context::Outside),
        }
        Ok(())
    }

    fn empty(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        let local = element.local_name();
        let core = is_spreadsheetml_name(namespace, element.name(), local.as_ref());
        if self.stack.is_empty() {
            return Err(invalid("worksheet root cannot be empty"));
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Worksheet, true, b"phoneticPr") => {
                self.add_properties(element, decoder, resolver)?;
            },
            (Context::PhoneticProperties, _, _) => {
                return Err(invalid("worksheet phoneticPr is a leaf element"));
            },
            _ => {},
        }
        Ok(())
    }

    fn end(&mut self, local: &[u8]) -> Result<()> {
        let context = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unexpected worksheet phoneticPr end element"))?;
        match context {
            Context::Worksheet if local == b"worksheet" => Ok(()),
            Context::PhoneticProperties if local == b"phoneticPr" => Ok(()),
            Context::Outside => Ok(()),
            _ => Err(invalid("mismatched worksheet phoneticPr end element")),
        }
    }

    fn add_properties(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_properties {
            return Err(invalid("duplicate worksheet phoneticPr element"));
        }
        self.seen_properties = true;
        self.properties = Some(parse_attributes(element, decoder, resolver)?);
        Ok(())
    }
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<WorksheetPhoneticProperties> {
    let mut font_id = None;
    let mut phonetic_type = None;
    let mut alignment = None;
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!(
                "unknown namespaced phoneticPr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"fontId" => set_once(&mut font_id, parse_font_id(&value)?, "fontId")?,
            b"type" => set_once(
                &mut phonetic_type,
                WorksheetPhoneticType::parse(&value)?,
                "type",
            )?,
            b"alignment" => set_once(
                &mut alignment,
                WorksheetPhoneticAlignment::parse(&value)?,
                "alignment",
            )?,
            name => {
                return Err(invalid(format!(
                    "unknown phoneticPr attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        }
    }
    Ok(WorksheetPhoneticProperties {
        font_id: font_id.ok_or_else(|| invalid("worksheet phoneticPr requires fontId"))?,
        phonetic_type: phonetic_type.unwrap_or(WorksheetPhoneticType::FullWidthKatakana),
        alignment: alignment.unwrap_or(WorksheetPhoneticAlignment::Left),
    })
}

fn parse_font_id(value: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid worksheet phoneticPr fontId '{value}'")))
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(format!("duplicate phoneticPr {name} attribute")));
    }
    *slot = Some(value);
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn reject_unsafe_event(event: &Event<'_>) -> Result<()> {
    if matches!(event, Event::DocType(_) | Event::PI(_)) {
        return Err(invalid("DTD and processing instructions are rejected"));
    }
    if let Event::GeneralRef(reference) = event {
        let name = reference.decode().map_err(xml_error)?;
        if !matches!(name.as_ref(), "amp" | "lt" | "gt" | "apos" | "quot") && !name.starts_with('#')
        {
            return Err(invalid("custom XML entities are rejected"));
        }
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(format!(
        "invalid worksheet phoneticPr XML: {error}"
    )))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetPhoneticProperties>> {
        parse_worksheet_phonetic_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_all_values_and_effective_defaults() {
        let defaults = parse(r#"<phoneticPr fontId="4294967295"/>"#)
            .unwrap()
            .unwrap();
        assert_eq!(defaults.font_id(), u32::MAX);
        assert_eq!(
            defaults.phonetic_type(),
            WorksheetPhoneticType::FullWidthKatakana
        );
        assert_eq!(defaults.alignment(), WorksheetPhoneticAlignment::Left);

        for (lexical, expected) in [
            (
                "halfwidthKatakana",
                WorksheetPhoneticType::HalfWidthKatakana,
            ),
            (
                "fullwidthKatakana",
                WorksheetPhoneticType::FullWidthKatakana,
            ),
            ("Hiragana", WorksheetPhoneticType::Hiragana),
            ("noConversion", WorksheetPhoneticType::NoConversion),
        ] {
            let value = parse(&format!(r#"<phoneticPr fontId="3" type="{lexical}"/>"#))
                .unwrap()
                .unwrap();
            assert_eq!(value.phonetic_type(), expected);
        }
        for (lexical, expected) in [
            ("noControl", WorksheetPhoneticAlignment::NoControl),
            ("left", WorksheetPhoneticAlignment::Left),
            ("center", WorksheetPhoneticAlignment::Center),
            ("distributed", WorksheetPhoneticAlignment::Distributed),
        ] {
            let value = parse(&format!(
                r#"<phoneticPr fontId="3" alignment="{lexical}"/>"#
            ))
            .unwrap()
            .unwrap();
            assert_eq!(value.alignment(), expected);
        }
        assert!(parse("").unwrap().is_none());
    }

    #[test]
    fn supports_strict_mce_and_exact_scoping() {
        let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><phoneticPr fontId="9" type="Hiragana" alignment="center"/></worksheet>"#;
        let value = parse_worksheet_phonetic_properties(strict)
            .unwrap()
            .unwrap();
        assert_eq!(value.font_id(), 9);
        assert_eq!(value.phonetic_type(), WorksheetPhoneticType::Hiragana);
        assert_eq!(value.alignment(), WorksheetPhoneticAlignment::Center);

        let mce = format!(
            concat!(
                r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x">"#,
                r#"<mc:AlternateContent><mc:Choice Requires="x"><x:phoneticPr fontId="8"/>"#,
                r#"</mc:Choice><mc:Fallback><phoneticPr fontId="7" x:future="1"/>"#,
                r#"</mc:Fallback></mc:AlternateContent></worksheet>"#,
            ),
            NS
        );
        assert_eq!(
            parse_worksheet_phonetic_properties(mce.as_bytes())
                .unwrap()
                .unwrap()
                .font_id(),
            7
        );
        assert!(
            parse("<wrapper><phoneticPr fontId=\"1\"/></wrapper>")
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn rejects_invalid_missing_duplicate_spoofed_and_non_leaf_values() {
        for child in [
            r#"<phoneticPr/>"#,
            r#"<phoneticPr fontId="-1"/>"#,
            r#"<phoneticPr fontId="4294967296"/>"#,
            r#"<phoneticPr fontId="1" type="hiragana"/>"#,
            r#"<phoneticPr fontId="1" type="unknown"/>"#,
            r#"<phoneticPr fontId="1" alignment="right"/>"#,
            r#"<phoneticPr fontId="1" mystery="x"/>"#,
            r#"<phoneticPr xmlns:x="urn:x" fontId="1" x:type="Hiragana"/>"#,
            r#"<phoneticPr fontId="1"><child/></phoneticPr>"#,
            r#"<phoneticPr fontId="1">text</phoneticPr>"#,
            r#"<phoneticPr fontId="1"/><phoneticPr fontId="2"/>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
    }

    #[test]
    fn rejects_malformed_document_boundaries_and_excessive_depth() {
        for xml in [
            format!(r#"<worksheet xmlns="{NS}"/><worksheet xmlns="{NS}"/>"#),
            format!(r#"text<worksheet xmlns="{NS}"></worksheet>"#),
            format!(r#"<worksheet xmlns="{NS}">text</worksheet>"#),
            format!(r#"<worksheet xmlns="{NS}"></worksheet>tail"#),
            format!(r#"<worksheet xmlns="{NS}"><![CDATA[data]]></worksheet>"#),
            format!(r#"<worksheet xmlns="{NS}"><phoneticPr></worksheet>"#),
        ] {
            assert!(
                parse_worksheet_phonetic_properties(xml.as_bytes()).is_err(),
                "expected rejection for {xml}"
            );
        }

        let mut xml = format!(r#"<worksheet xmlns="{NS}">"#);
        for _ in 0..MAX_DEPTH {
            xml.push_str("<extension>");
        }
        for _ in 0..MAX_DEPTH {
            xml.push_str("</extension>");
        }
        xml.push_str("</worksheet>");
        assert!(parse_worksheet_phonetic_properties(xml.as_bytes()).is_err());
    }

    fn fixture(bytes: &[u8]) -> WorksheetPhoneticProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        parse_worksheet_phonetic_properties(part.blob())
            .unwrap()
            .unwrap()
    }

    #[test]
    fn reads_poi_and_libreoffice_fixtures() {
        let preserve_attributes = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/48962.xlsx"
        )));
        assert_eq!(preserve_attributes.font_id(), 3);
        assert_eq!(
            preserve_attributes.phonetic_type(),
            WorksheetPhoneticType::NoConversion
        );

        let poi = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/54071.xlsx"
        )));
        assert_eq!(poi.font_id(), 1);
        assert_eq!(poi.phonetic_type(), WorksheetPhoneticType::NoConversion);
        assert_eq!(poi.alignment(), WorksheetPhoneticAlignment::Left);

        let libreoffice = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf97598_scenarios.xlsx"
        )));
        assert_eq!(libreoffice.font_id(), 1);
        assert_eq!(
            libreoffice.phonetic_type(),
            WorksheetPhoneticType::NoConversion
        );
        assert_eq!(libreoffice.alignment(), WorksheetPhoneticAlignment::Left);
    }
}
