//! Namespace-aware, bounded retained chart-content reader.

use crate::namespace::{CHARTNS, OFFICENS};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_ELEMENTS: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 16 * 1_048_576;

/// A recognized local element in the standard chart vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    Chart,
    Title,
    Subtitle,
    Footer,
    Legend,
    PlotArea,
    Wall,
    Floor,
    Axis,
    Categories,
    Grid,
    Series,
    Domain,
    DataPoint,
    DataLabel,
    MeanValue,
    ErrorIndicator,
    RegressionCurve,
    Equation,
    StockGainMarker,
    StockLossMarker,
    StockRangeLine,
    SymbolImage,
    LabelSeparator,
    /// An element outside the standard chart namespace or a future element.
    Other,
}

/// A decoded XML attribute identified by its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
    value_namespace_uri: Option<String>,
}

impl Attribute {
    #[must_use]
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    #[must_use]
    pub fn value(&self) -> &str {
        &self.value
    }
}

/// An ordered, retained element in a chart content subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Element {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<Attribute>,
    text: String,
    children: Vec<Element>,
}

impl Element {
    #[must_use]
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    #[must_use]
    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    #[must_use]
    pub fn kind(&self) -> Kind {
        if self.namespace_uri() != Some(CHARTNS) {
            return Kind::Other;
        }
        match self.local_name.as_str() {
            "chart" => Kind::Chart,
            "title" => Kind::Title,
            "subtitle" => Kind::Subtitle,
            "footer" => Kind::Footer,
            "legend" => Kind::Legend,
            "plot-area" => Kind::PlotArea,
            "wall" => Kind::Wall,
            "floor" => Kind::Floor,
            "axis" => Kind::Axis,
            "categories" => Kind::Categories,
            "grid" => Kind::Grid,
            "series" => Kind::Series,
            "domain" => Kind::Domain,
            "data-point" => Kind::DataPoint,
            "data-label" => Kind::DataLabel,
            "mean-value" => Kind::MeanValue,
            "error-indicator" => Kind::ErrorIndicator,
            "regression-curve" => Kind::RegressionCurve,
            "equation" => Kind::Equation,
            "stock-gain-marker" => Kind::StockGainMarker,
            "stock-loss-marker" => Kind::StockLossMarker,
            "stock-range-line" => Kind::StockRangeLine,
            "symbol-image" => Kind::SymbolImage,
            "label-separator" => Kind::LabelSeparator,
            _ => Kind::Other,
        }
    }

    #[must_use]
    pub fn attributes(&self) -> &[Attribute] {
        &self.attributes
    }

    #[must_use]
    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(Attribute::value)
    }

    /// Decode the element's `chart:class` value as a typed namespaced token.
    ///
    /// The parser retains the exact QName spelling and resolves its prefix in
    /// the producer's namespace context, so aliases are not normalized.
    pub fn chart_class(&self) -> Result<super::ChartClass> {
        let attribute = self
            .attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri.as_deref() == Some(CHARTNS)
                    && attribute.local_name == "class"
            })
            .ok_or_else(|| invalid_error("chart:chart requires chart:class"))?;
        super::ChartClass::parse(&attribute.value, attribute.value_namespace_uri.as_deref())
    }

    /// Return direct character content, excluding descendant text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    #[must_use]
    pub fn children(&self) -> &[Element] {
        &self.children
    }

    pub fn children_of_kind(&self, kind: Kind) -> impl Iterator<Item = &Element> {
        self.children
            .iter()
            .filter(move |child| child.kind() == kind)
    }

    /// Compose character content from this element and all descendants.
    #[must_use]
    pub fn all_text(&self) -> String {
        let mut output = String::new();
        let mut stack = vec![self];
        while let Some(node) = stack.pop() {
            output.push_str(&node.text);
            stack.extend(node.children.iter().rev());
        }
        output
    }
}

/// Read and retain the standard chart subtree from an ODF `content.xml` part.
///
/// Namespace prefixes are resolved before validation. Unknown elements and
/// attributes are retained as inert expanded names, so a family owner can
/// inspect or copy vendor extensions without interpreting them.
///
/// # Errors
///
/// Returns an invalid-format error when the XML is malformed, the required
/// ODF chart structure is absent, or a configured size/depth limit is exceeded.
pub fn read(xml: &str) -> Result<Element> {
    if xml.len() > MAX_CONTENT_BYTES {
        return invalid("chart content exceeds 256 MiB");
    }

    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut office_chart_seen = false;
    let mut body_depth = None;
    let mut office_chart_depth = None;
    let mut chart_depth = None;
    let mut chart_complete = None;
    let mut stack = Vec::new();
    let mut element_count = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid_error(format!("invalid chart XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || namespace_uri.as_deref() != Some(OFFICENS)
                        || local != "document-content"
                    {
                        return invalid("chart content must have one office:document-content root");
                    }
                    root_seen = true;
                } else if namespace_uri.as_deref() == Some(OFFICENS) && local == "body" {
                    if depth != 1 || body_seen || body_depth.is_some() {
                        return invalid("misplaced or duplicate office:body");
                    }
                    body_seen = true;
                    body_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(OFFICENS) && local == "chart" {
                    if depth != 2
                        || body_depth != Some(2)
                        || office_chart_seen
                        || office_chart_depth.is_some()
                    {
                        return invalid("misplaced or duplicate office:chart");
                    }
                    office_chart_seen = true;
                    office_chart_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(CHARTNS) && local == "chart" {
                    if depth != 3
                        || office_chart_depth != Some(3)
                        || chart_depth.is_some()
                        || chart_complete.is_some()
                    {
                        return invalid("misplaced or duplicate chart:chart");
                    }
                    chart_depth = Some(depth + 1);
                }
                if chart_depth.is_some() {
                    push_node(
                        &reader,
                        element,
                        namespace_uri,
                        local,
                        &mut stack,
                        &mut element_count,
                    )?;
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid_error("chart XML nesting overflow"))?;
                if depth > MAX_DEPTH || stack.len() > MAX_DEPTH {
                    return invalid("chart element nesting exceeds 128 levels");
                }
            },
            Event::Empty(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                if depth == 0 {
                    return invalid("chart content root cannot be empty");
                }
                if namespace_uri.as_deref() == Some(CHARTNS) && local == "chart" {
                    return invalid("chart:chart cannot be empty");
                }
                if namespace_uri.as_deref() == Some(OFFICENS) && local == "body" {
                    if depth != 1 || body_seen {
                        return invalid("misplaced or duplicate office:body");
                    }
                    body_seen = true;
                } else if namespace_uri.as_deref() == Some(OFFICENS) && local == "chart" {
                    if depth != 2 || body_depth != Some(2) || office_chart_seen {
                        return invalid("misplaced or duplicate office:chart");
                    }
                    office_chart_seen = true;
                }
                if chart_depth.is_some() {
                    let node =
                        make_node(&reader, element, namespace_uri, local, &mut element_count)?;
                    stack
                        .last_mut()
                        .ok_or_else(|| invalid_error("chart root is not active"))?
                        .children
                        .push(node);
                }
            },
            Event::End(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid_error("unexpected chart XML closing tag"))?;
                if chart_depth.is_some() {
                    let node = stack
                        .pop()
                        .ok_or_else(|| invalid_error("chart node stack underflow"))?;
                    if stack.is_empty() {
                        chart_complete = Some(node);
                        chart_depth = None;
                    } else {
                        stack
                            .last_mut()
                            .ok_or_else(|| invalid_error("chart parent is missing"))?
                            .children
                            .push(node);
                    }
                }
                if namespace_uri.as_deref() == Some(OFFICENS) && local == "chart" && depth == 2 {
                    office_chart_depth = None;
                } else if namespace_uri.as_deref() == Some(OFFICENS)
                    && local == "body"
                    && depth == 1
                {
                    body_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid_error(format!("invalid chart text: {error}")))?;
                append_text(
                    stack
                        .last_mut()
                        .ok_or_else(|| invalid_error("chart node is missing"))?,
                    &value,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| invalid_error(format!("invalid chart CDATA: {error}")))?;
                append_text(
                    stack
                        .last_mut()
                        .ok_or_else(|| invalid_error("chart node is missing"))?,
                    &value,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                let value = decode_reference(reference)?;
                append_text(
                    stack
                        .last_mut()
                        .ok_or_else(|| invalid_error("chart node is missing"))?,
                    &value,
                )?;
            },
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return invalid("text is not allowed outside the chart content root");
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return invalid("content is not allowed outside the chart content root");
            },
            Event::Eof => break,
            Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::Comment(_)
            | Event::Text(_)
            | Event::CData(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || !office_chart_seen
        || body_depth.is_some()
        || office_chart_depth.is_some()
        || chart_depth.is_some()
        || !stack.is_empty()
    {
        return invalid("incomplete standalone chart structure");
    }
    chart_complete.ok_or_else(|| invalid_error("standalone chart has no chart:chart"))
}

fn push_node(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace_uri: Option<String>,
    local_name: String,
    stack: &mut Vec<Element>,
    element_count: &mut usize,
) -> Result<()> {
    stack.push(make_node(
        reader,
        element,
        namespace_uri,
        local_name,
        element_count,
    )?);
    Ok(())
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    resolved_namespace_uri: Option<String>,
    local_name: String,
    element_count: &mut usize,
) -> Result<Element> {
    *element_count = element_count
        .checked_add(1)
        .ok_or_else(|| invalid_error("chart element count overflow"))?;
    if *element_count > MAX_ELEMENTS {
        return invalid("chart exceeds 65536 elements");
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return invalid("chart element exceeds 256 attributes");
    }

    let mut attributes = Vec::new();
    for raw_attribute in element.attributes() {
        let attribute = raw_attribute
            .map_err(|error| invalid_error(format!("invalid chart attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let attribute_name = decode_name(local.as_ref(), "attribute")?;
        if attributes.iter().any(|existing: &Attribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == attribute_name
        }) {
            return invalid(format!(
                "duplicate expanded chart attribute '{attribute_name}'"
            ));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid_error(format!("invalid chart attribute value: {error}")))?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return invalid("chart attribute exceeds 1 MiB");
        }
        let value_namespace_uri = if namespace_uri.as_deref() == Some(CHARTNS)
            && attribute_name == "class"
        {
            let (value_namespace, _) = reader.resolver().resolve_element(QName(value.as_bytes()));
            match value_namespace {
                ResolveResult::Unbound => None,
                ResolveResult::Bound(Namespace(uri)) => {
                    Some(decode_name(uri, "chart class namespace URI")?)
                },
                ResolveResult::Unknown(prefix) => {
                    return invalid(format!(
                        "unknown chart class namespace prefix '{}'",
                        String::from_utf8_lossy(&prefix)
                    ));
                },
            }
        } else {
            None
        };
        attributes.push(Attribute {
            namespace_uri,
            local_name: attribute_name,
            value,
            value_namespace_uri,
        });
    }

    Ok(Element {
        namespace_uri: resolved_namespace_uri,
        local_name,
        attributes,
        text: String::new(),
        children: Vec::new(),
    })
}

fn append_text(node: &mut Element, value: &str) -> Result<()> {
    if node.text.len().saturating_add(value.len()) > MAX_TEXT_BYTES {
        return invalid("chart text node exceeds 16 MiB");
    }
    node.text.push_str(value);
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_name(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unknown chart namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        )),
    }
}

fn decode_name(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|error| invalid_error(format!("non-UTF-8 chart {kind}: {error}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| invalid_error(format!("invalid chart character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| invalid_error(format!("invalid chart entity reference: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => invalid(format!("unsupported chart entity reference '&{name};'")),
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const CHART: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn content(body: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:c="{CHART}" xmlns:t="{TABLE}" xmlns:x="{TEXT}"><o:body><o:chart><c:chart c:class="c:bar">{body}</c:chart></o:chart></o:body></o:document-content>"#
        )
    }

    #[test]
    fn retains_namespace_aware_extension_content() {
        let xml = content(
            r#"<c:title><x:p>Revenue &amp; margin</x:p></c:title><vendor:extension xmlns:vendor="urn:vendor:chart" vendor:flag="yes"><vendor:value><![CDATA[opaque <value>]]></vendor:value></vendor:extension>"#,
        );
        let chart = read(&xml).unwrap();
        assert_eq!(chart.kind(), Kind::Chart);
        assert_eq!(chart.children_of_kind(Kind::Title).count(), 1);
        let extension = chart.children().last().unwrap();
        assert_eq!(extension.namespace_uri(), Some("urn:vendor:chart"));
        assert_eq!(extension.all_text(), "opaque <value>");
        assert_eq!(
            extension.attribute(Some("urn:vendor:chart"), "flag"),
            Some("yes")
        );
    }

    #[test]
    fn resolves_namespace_aliases_and_rejects_expanded_duplicates() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:c="{CHART}" xmlns:other="{CHART}"><o:body><o:chart><c:chart c:class="c:bar" other:class="c:line"><c:plot-area/></c:chart></o:chart></o:body></o:document-content>"#
        );
        assert!(read(&xml).is_err());

        let valid = content(
            r#"<c:plot-area table:cell-range-address="Data.A1:C4" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"/>"#,
        );
        assert_eq!(
            read(&valid)
                .unwrap()
                .children_of_kind(Kind::PlotArea)
                .count(),
            1
        );
    }

    #[test]
    fn enforces_structure_depth_and_entity_rules() {
        let missing = format!(
            r#"<o:document-content xmlns:o="{OFFICE}"><o:body><o:chart/></o:body></o:document-content>"#
        );
        assert!(read(&missing).is_err());

        let nested = "<c:series>".repeat(MAX_DEPTH + 1) + &"</c:series>".repeat(MAX_DEPTH + 1);
        assert!(read(&content(&nested)).is_err());

        let unsupported = content(r"<c:title>&unknown;</c:title>");
        assert!(read(&unsupported).is_err());
    }
}
