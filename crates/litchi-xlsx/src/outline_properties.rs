//! Immutable XLSX worksheet outline-properties read model.

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

/// Effective grouping and outline-display policy from `sheetPr/outlinePr`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OutlineProperties {
    apply_styles: bool,
    summary_below: bool,
    summary_right: bool,
    show_outline_symbols: bool,
}

impl OutlineProperties {
    pub fn apply_styles(&self) -> bool {
        self.apply_styles
    }
    pub fn summary_below(&self) -> bool {
        self.summary_below
    }
    pub fn summary_right(&self) -> bool {
        self.summary_right
    }
    /// Stored sheet-level preference. A sheet-view value overrides this on conflict.
    pub fn show_outline_symbols(&self) -> bool {
        self.show_outline_symbols
    }
}

#[derive(Default)]
struct Builder {
    apply_styles: Option<bool>,
    summary_below: Option<bool>,
    summary_right: Option<bool>,
    show_outline_symbols: Option<bool>,
}

impl Builder {
    fn finish(self) -> OutlineProperties {
        OutlineProperties {
            apply_styles: self.apply_styles.unwrap_or(false),
            summary_below: self.summary_below.unwrap_or(true),
            summary_right: self.summary_right.unwrap_or(true),
            show_outline_symbols: self.show_outline_symbols.unwrap_or(true),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Context {
    Outside,
    Worksheet,
    SheetProperties,
    OutlineProperties,
}

struct Parser {
    stack: Vec<Context>,
    properties: Option<OutlineProperties>,
    seen_sheet_properties: bool,
    seen_outline_properties: bool,
}

/// Parse the worksheet's exact `worksheet/sheetPr/outlinePr` child path.
// Text/CData arms keep `?`-bearing whitespace checks out of guards; guards cannot use `?`.
#[allow(clippy::collapsible_match)]
pub fn parse_outline_properties(xml: &[u8]) -> Result<Option<OutlineProperties>> {
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
        seen_sheet_properties: false,
        seen_outline_properties: false,
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
                if matches!(
                    parser.parent(),
                    Context::SheetProperties | Context::OutlineProperties
                ) =>
            {
                if !text.decode().map_err(xml_error)?.trim().is_empty() {
                    return Err(invalid("unexpected text in worksheet outline properties"));
                }
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
            Event::CData(_)
                if matches!(
                    parser.parent(),
                    Context::SheetProperties | Context::OutlineProperties
                ) =>
            {
                return Err(invalid("unexpected CDATA in worksheet outline properties"));
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
                    || matches!(
                        parser.parent(),
                        Context::SheetProperties | Context::OutlineProperties
                    )
                {
                    return Err(invalid(
                        "worksheet outline properties cannot contain entity text",
                    ));
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
        return Err(invalid("unterminated worksheet outline XML"));
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
            return Err(invalid("outlinePr parser requires a worksheet root"));
        }
        match (self.parent(), core, local.as_ref()) {
            (Context::Outside, true, b"worksheet") => self.stack.push(Context::Worksheet),
            (Context::Worksheet, true, b"sheetPr") => {
                self.begin_sheet_properties(element)?;
                self.stack.push(Context::SheetProperties);
            },
            (Context::SheetProperties, true, b"outlinePr") => {
                self.add_outline_properties(element, decoder, resolver)?;
                self.stack.push(Context::OutlineProperties);
            },
            (Context::OutlineProperties, _, _) => {
                return Err(invalid("outlinePr is a leaf element"));
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
            (Context::Worksheet, true, b"sheetPr") => self.begin_sheet_properties(element)?,
            (Context::SheetProperties, true, b"outlinePr") => {
                self.add_outline_properties(element, decoder, resolver)?;
            },
            (Context::OutlineProperties, _, _) => {
                return Err(invalid("outlinePr is a leaf element"));
            },
            _ => {},
        }
        Ok(())
    }

    fn end(&mut self, local: &[u8]) -> Result<()> {
        let context = self
            .stack
            .pop()
            .ok_or_else(|| invalid("unexpected worksheet outline end element"))?;
        match context {
            Context::Worksheet if local == b"worksheet" => Ok(()),
            Context::SheetProperties if local == b"sheetPr" => Ok(()),
            Context::OutlineProperties if local == b"outlinePr" => Ok(()),
            Context::Outside => Ok(()),
            _ => Err(invalid("mismatched worksheet outline end element")),
        }
    }

    fn begin_sheet_properties(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self.seen_sheet_properties {
            return Err(invalid("duplicate worksheet sheetPr element"));
        }
        validate_attribute_syntax(element)?;
        self.seen_sheet_properties = true;
        Ok(())
    }

    fn add_outline_properties(
        &mut self,
        element: &BytesStart<'_>,
        decoder: Decoder,
        resolver: &NamespaceResolver,
    ) -> Result<()> {
        if self.seen_outline_properties {
            return Err(invalid("duplicate worksheet outlinePr element"));
        }
        self.seen_outline_properties = true;
        self.properties = Some(parse_attributes(element, decoder, resolver)?.finish());
        Ok(())
    }
}

fn parse_attributes(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<Builder> {
    let mut builder = Builder::default();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(xml_error)?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = resolver.resolve_attribute(attribute.key);
        if !matches!(namespace, ResolveResult::Unbound) {
            return Err(invalid(format!(
                "unknown namespaced outlinePr attribute '{}'",
                String::from_utf8_lossy(attribute.key.as_ref())
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(xml_error)?;
        match local.as_ref() {
            b"applyStyles" => set_once(
                &mut builder.apply_styles,
                parse_bool(&value, "applyStyles")?,
                "applyStyles",
            )?,
            b"summaryBelow" => set_once(
                &mut builder.summary_below,
                parse_bool(&value, "summaryBelow")?,
                "summaryBelow",
            )?,
            b"summaryRight" => set_once(
                &mut builder.summary_right,
                parse_bool(&value, "summaryRight")?,
                "summaryRight",
            )?,
            b"showOutlineSymbols" => set_once(
                &mut builder.show_outline_symbols,
                parse_bool(&value, "showOutlineSymbols")?,
                "showOutlineSymbols",
            )?,
            name => {
                return Err(invalid(format!(
                    "unknown outlinePr attribute '{}'",
                    String::from_utf8_lossy(name)
                )));
            },
        }
    }
    Ok(builder)
}

fn validate_attribute_syntax(element: &BytesStart<'_>) -> Result<()> {
    for attribute in element.attributes().with_checks(true) {
        attribute.map_err(xml_error)?;
    }
    Ok(())
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.is_some() {
        return Err(invalid(format!("duplicate outlinePr {name} attribute")));
    }
    *slot = Some(value);
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "1" | "true" => Ok(true),
        "0" | "false" => Ok(false),
        _ => Err(invalid(format!(
            "invalid outlinePr {name} boolean '{value}'"
        ))),
    }
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
    invalid(format!("invalid worksheet outlinePr XML: {error}"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<OutlineProperties>> {
        parse_outline_properties(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_all_attributes_and_effective_defaults() {
        let value = parse(concat!(
            r#"<sheetPr><outlinePr applyStyles="1" summaryBelow="0" "#,
            r#"summaryRight="false" showOutlineSymbols="0"/></sheetPr>"#
        ))
        .unwrap()
        .unwrap();
        assert!(value.apply_styles());
        assert!(!value.summary_below());
        assert!(!value.summary_right());
        assert!(!value.show_outline_symbols());

        let defaults = parse("<sheetPr><outlinePr/></sheetPr>").unwrap().unwrap();
        assert!(!defaults.apply_styles());
        assert!(defaults.summary_below());
        assert!(defaults.summary_right());
        assert!(defaults.show_outline_symbols());
        assert!(parse("<sheetPr/>").unwrap().is_none());
    }

    #[test]
    fn supports_strict_mce_and_exact_scoping() {
        let strict = br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetPr><outlinePr summaryRight="0"/></sheetPr></worksheet>"#;
        assert!(
            !parse_outline_properties(strict)
                .unwrap()
                .unwrap()
                .summary_right()
        );
        let mce = format!(
            concat!(
                r#"<worksheet xmlns="{}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x">"#,
                r#"<sheetPr codeName="ignored"><tabColor rgb="FF00FF00"/><mc:AlternateContent>"#,
                r#"<mc:Choice Requires="x"><x:outlinePr/></mc:Choice><mc:Fallback><outlinePr summaryBelow="0"/>"#,
                r#"</mc:Fallback></mc:AlternateContent><pageSetUpPr fitToPage="1"/></sheetPr></worksheet>"#,
            ),
            NS
        );
        assert!(
            !parse_outline_properties(mce.as_bytes())
                .unwrap()
                .unwrap()
                .summary_below()
        );
        assert!(
            parse("<wrapper><sheetPr><outlinePr/></sheetPr></wrapper>")
                .unwrap()
                .is_none()
        );
        assert!(
            parse("<sheetPr><wrapper><outlinePr/></wrapper></sheetPr>")
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn rejects_invalid_values_duplicates_spoofing_and_children() {
        for child in [
            r#"<sheetPr><outlinePr summaryBelow="yes"/></sheetPr>"#,
            r#"<sheetPr><outlinePr mystery="1"/></sheetPr>"#,
            r#"<sheetPr xmlns:x="urn:x"><outlinePr x:summaryBelow="0"/></sheetPr>"#,
            r#"<sheetPr><outlinePr><child/></outlinePr></sheetPr>"#,
            r#"<sheetPr><outlinePr/><outlinePr/></sheetPr>"#,
            r#"<sheetPr/><sheetPr/>"#,
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
            format!(r#"<worksheet xmlns="{NS}"><sheetPr></worksheet>"#),
        ] {
            assert!(
                parse_outline_properties(xml.as_bytes()).is_err(),
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
        assert!(parse_outline_properties(xml.as_bytes()).is_err());
    }

    fn fixture(bytes: &[u8]) -> OutlineProperties {
        let package = OpcPackage::from_bytes(bytes).unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        parse_outline_properties(part.blob()).unwrap().unwrap()
    }

    #[test]
    fn reads_poi_and_libreoffice_outline_fixtures() {
        let poi = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/poi/test-data/spreadsheet/66365.xlsx"
        )));
        assert!(!poi.summary_below());
        assert!(!poi.summary_right());

        let subtotal = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/subtotal-above.xlsx"
        )));
        assert!(!subtotal.summary_below());
        assert!(subtotal.summary_right());

        let defaults = fixture(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/totalsRowShown.xlsx"
        )));
        assert!(!defaults.apply_styles());
        assert!(defaults.summary_below());
        assert!(defaults.summary_right());
        assert!(defaults.show_outline_symbols());
    }
}
