//! Namespace-aware read-only access to standalone OpenDocument charts.

use crate::{OdfMetadata, OpenDocumentFamily, OpenDocumentPackage};
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CHART_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";

/// A recognized element in the standard ODF chart vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum ChartElementKind {
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
    /// An element outside the standard chart namespace or a future chart element.
    Other,
}

/// One decoded XML attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartAttribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl ChartAttribute {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

/// An ordered element in the standalone chart's complete XML subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChartElement {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<ChartAttribute>,
    text: String,
    children: Vec<ChartElement>,
}

impl ChartElement {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn kind(&self) -> ChartElementKind {
        if self.namespace_uri() != Some(CHART_NAMESPACE) {
            return ChartElementKind::Other;
        }
        match self.local_name.as_str() {
            "chart" => ChartElementKind::Chart,
            "title" => ChartElementKind::Title,
            "subtitle" => ChartElementKind::Subtitle,
            "footer" => ChartElementKind::Footer,
            "legend" => ChartElementKind::Legend,
            "plot-area" => ChartElementKind::PlotArea,
            "wall" => ChartElementKind::Wall,
            "floor" => ChartElementKind::Floor,
            "axis" => ChartElementKind::Axis,
            "categories" => ChartElementKind::Categories,
            "grid" => ChartElementKind::Grid,
            "series" => ChartElementKind::Series,
            "domain" => ChartElementKind::Domain,
            "data-point" => ChartElementKind::DataPoint,
            "data-label" => ChartElementKind::DataLabel,
            "mean-value" => ChartElementKind::MeanValue,
            "error-indicator" => ChartElementKind::ErrorIndicator,
            "regression-curve" => ChartElementKind::RegressionCurve,
            "equation" => ChartElementKind::Equation,
            "stock-gain-marker" => ChartElementKind::StockGainMarker,
            "stock-loss-marker" => ChartElementKind::StockLossMarker,
            "stock-range-line" => ChartElementKind::StockRangeLine,
            "symbol-image" => ChartElementKind::SymbolImage,
            "label-separator" => ChartElementKind::LabelSeparator,
            _ => ChartElementKind::Other,
        }
    }

    pub fn attributes(&self) -> &[ChartAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(ChartAttribute::value)
    }

    /// Return direct character content, excluding descendant text.
    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn children(&self) -> &[ChartElement] {
        &self.children
    }

    pub fn children_of_kind(&self, kind: ChartElementKind) -> impl Iterator<Item = &ChartElement> {
        self.children
            .iter()
            .filter(move |child| child.kind() == kind)
    }

    /// Compose character content from this element and all descendants.
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

/// A validated standalone OpenDocument chart or chart template.
pub struct ChartDocument {
    package: OpenDocumentPackage,
    chart: ChartElement,
}

impl ChartDocument {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OpenDocumentPackage::from_bytes(bytes)?;
        if package.family() != OpenDocumentFamily::Chart {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument chart: MIME type is '{}'",
                package.mimetype()
            )));
        }
        let chart = parse_chart_content(&package.content_xml()?)?;
        Ok(Self { package, chart })
    }

    pub fn is_template(&self) -> bool {
        self.package.is_template()
    }

    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    pub fn chart(&self) -> &ChartElement {
        &self.chart
    }

    pub fn text(&self) -> String {
        self.chart.all_text()
    }

    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    pub fn odf_metadata(&self) -> Result<Option<OdfMetadata>> {
        self.package.odf_metadata()
    }

    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn parse_chart_content(xml: &str) -> Result<ChartElement> {
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
    let mut node_count = 0usize;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid chart XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || namespace_uri.as_deref() != Some(OFFICE_NAMESPACE)
                        || local != "document-content"
                    {
                        return Err(Error::InvalidFormat(
                            "chart content must have one office:document-content root".to_string(),
                        ));
                    }
                    root_seen = true;
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE) && local == "body" {
                    if depth != 1 || body_seen || body_depth.is_some() {
                        return Err(Error::InvalidFormat(
                            "misplaced or duplicate office:body".to_string(),
                        ));
                    }
                    body_seen = true;
                    body_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE) && local == "chart" {
                    if depth != 2
                        || body_depth != Some(2)
                        || office_chart_seen
                        || office_chart_depth.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "misplaced or duplicate office:chart".to_string(),
                        ));
                    }
                    office_chart_seen = true;
                    office_chart_depth = Some(depth + 1);
                } else if namespace_uri.as_deref() == Some(CHART_NAMESPACE) && local == "chart" {
                    if depth != 3
                        || office_chart_depth != Some(3)
                        || chart_depth.is_some()
                        || chart_complete.is_some()
                    {
                        return Err(Error::InvalidFormat(
                            "misplaced or duplicate chart:chart".to_string(),
                        ));
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
                        &mut node_count,
                    )?;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("chart XML nesting overflow".to_string())
                })?;
                if stack.len() > 128 {
                    return Err(Error::InvalidFormat(
                        "chart element nesting exceeds 128 levels".to_string(),
                    ));
                }
            },
            Event::Empty(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "chart content root cannot be empty".to_string(),
                    ));
                }
                if namespace_uri.as_deref() == Some(CHART_NAMESPACE) && local == "chart" {
                    return Err(Error::InvalidFormat(
                        "chart:chart cannot be empty".to_string(),
                    ));
                }
                if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE) && local == "body" {
                    if depth != 1 || body_seen {
                        return Err(Error::InvalidFormat(
                            "misplaced or duplicate office:body".to_string(),
                        ));
                    }
                    body_seen = true;
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE) && local == "chart" {
                    if depth != 2 || body_depth != Some(2) || office_chart_seen {
                        return Err(Error::InvalidFormat(
                            "misplaced or duplicate office:chart".to_string(),
                        ));
                    }
                    office_chart_seen = true;
                }
                if chart_depth.is_some() {
                    let node = make_node(&reader, element, namespace_uri, local, &mut node_count)?;
                    stack
                        .last_mut()
                        .expect("chart root is active")
                        .children
                        .push(node);
                }
            },
            Event::End(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_name(element.local_name().as_ref(), "element")?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected chart XML closing tag".to_string())
                })?;
                if chart_depth.is_some() {
                    let node = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("chart node stack underflow".to_string())
                    })?;
                    if stack.is_empty() {
                        chart_complete = Some(node);
                        chart_depth = None;
                    } else {
                        stack.last_mut().expect("parent exists").children.push(node);
                    }
                }
                if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "chart"
                    && depth == 2
                {
                    office_chart_depth = None;
                } else if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
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
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid chart text: {error}"))
                })?;
                append_text(stack.last_mut().expect("node exists"), &value)?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid chart CDATA: {error}"))
                })?;
                append_text(stack.last_mut().expect("node exists"), &value)?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => {
                let value = decode_reference(reference)?;
                append_text(stack.last_mut().expect("node exists"), &value)?;
            },
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the chart content root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the chart content root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
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
        return Err(Error::InvalidFormat(
            "incomplete standalone chart structure".to_string(),
        ));
    }
    chart_complete
        .ok_or_else(|| Error::InvalidFormat("standalone chart has no chart:chart".to_string()))
}

fn push_node(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    namespace_uri: Option<String>,
    local_name: String,
    stack: &mut Vec<ChartElement>,
    node_count: &mut usize,
) -> Result<()> {
    let node = make_node(reader, element, namespace_uri, local_name, node_count)?;
    stack.push(node);
    Ok(())
}

fn make_node(
    reader: &NsReader<&[u8]>,
    element: &quick_xml::events::BytesStart<'_>,
    resolved_namespace_uri: Option<String>,
    local_name: String,
    node_count: &mut usize,
) -> Result<ChartElement> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("chart node count overflow".to_string()))?;
    if *node_count > 65_536 {
        return Err(Error::InvalidFormat(
            "chart exceeds 65536 elements".to_string(),
        ));
    }
    if element.attributes().count() > 256 {
        return Err(Error::InvalidFormat(
            "chart element exceeds 256 attributes".to_string(),
        ));
    }
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| Error::InvalidFormat(format!("invalid chart attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace_uri = namespace_uri(&namespace)?;
        let local_name = decode_name(local.as_ref(), "attribute")?;
        if attributes.iter().any(|existing: &ChartAttribute| {
            existing.namespace_uri == namespace_uri && existing.local_name == local_name
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded chart attribute '{local_name}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid chart attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > 1_048_576 {
            return Err(Error::InvalidFormat(
                "chart attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(ChartAttribute {
            namespace_uri,
            local_name,
            value,
        });
    }
    Ok(ChartElement {
        namespace_uri: resolved_namespace_uri,
        local_name,
        attributes,
        text: String::new(),
        children: Vec::new(),
    })
}

fn append_text(node: &mut ChartElement, value: &str) -> Result<()> {
    if node.text.len().saturating_add(value.len()) > 16 * 1_048_576 {
        return Err(Error::InvalidFormat(
            "chart text node exceeds 16 MiB".to_string(),
        ));
    }
    node.text.push_str(value);
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_name(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown chart namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode_name(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 chart {kind}")))
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid chart character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference.decode().map_err(|error| {
        Error::InvalidFormat(format!("invalid chart entity reference: {error}"))
    })?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported chart entity reference '&{name};'"
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn chart_xml() -> &'static str {
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
 xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:tb="urn:oasis:names:tc:opendocument:xmlns:table:1.0">
 <o:body><o:chart>
  <c:chart c:class="c:bar" s:width="12cm" s:height="8cm" c:column-mapping="1 2">
   <c:title tb:cell-range="Data.A1"><t:p>Revenue &amp; margin</t:p></c:title>
   <c:subtitle><t:p><![CDATA[2026 <plan>]]></t:p></c:subtitle>
   <c:legend c:legend-position="end"/>
   <c:plot-area tb:cell-range-address="Data.A1:C4" c:data-source-has-labels="both">
    <c:axis c:dimension="x" c:name="primary-x"><c:categories tb:cell-range-address="Data.A2:A4"/><c:grid c:class="major"/></c:axis>
    <c:series c:values-cell-range-address="Data.B2:B4" c:label-cell-address="Data.B1"><c:domain tb:cell-range-address="Data.A2:A4"/><c:data-point c:repeated="3"/></c:series>
    <c:series c:values-cell-range-address="Data.C2:C4"><c:mean-value/><c:regression-curve><c:equation c:display-equation="true"/></c:regression-curve><c:error-indicator/></c:series>
   </c:plot-area>
   <tb:table tb:name="Data"><tb:table-row><tb:table-cell o:value-type="string"><t:p>Revenue</t:p></tb:table-cell></tb:table-row></tb:table>
  </c:chart>
 </o:chart></o:body>
</o:document-content>"#
    }

    #[test]
    fn parses_complete_namespace_aware_chart_subtree_losslessly() {
        let bytes = package(constants::ODF_CHART, chart_xml());
        let document = ChartDocument::from_bytes(bytes.clone()).unwrap();
        assert!(!document.is_template());
        assert_eq!(document.chart().kind(), ChartElementKind::Chart);
        assert_eq!(
            document.chart().attribute(Some(CHART_NAMESPACE), "class"),
            Some("c:bar")
        );
        let title = document
            .chart()
            .children_of_kind(ChartElementKind::Title)
            .next()
            .unwrap();
        assert_eq!(title.all_text(), "Revenue & margin");
        let plot = document
            .chart()
            .children_of_kind(ChartElementKind::PlotArea)
            .next()
            .unwrap();
        assert_eq!(plot.children_of_kind(ChartElementKind::Axis).count(), 1);
        assert_eq!(plot.children_of_kind(ChartElementKind::Series).count(), 2);
        let table = document
            .chart()
            .children()
            .iter()
            .find(|child| child.local_name() == "table")
            .unwrap();
        assert!(table.all_text().contains("Revenue"));
        assert!(document.text().contains("2026 <plan>"));
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_chart_templates_and_readers() {
        let bytes = package(constants::ODF_CHART_TEMPLATE, chart_xml());
        let document = ChartDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert!(document.is_template());
        assert_eq!(document.into_bytes(), bytes);
    }

    #[test]
    fn rejects_other_families_and_invalid_chart_structure() {
        assert!(ChartDocument::from_bytes(package(constants::ODF_DRAWING, chart_xml())).is_err());
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:chart/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body><o:chart><c:chart>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body><c:chart/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body/><o:body><o:chart><c:chart><c:plot-area/></c:chart></o:chart></o:body></o:document-content>"#,
        ] {
            assert!(
                ChartDocument::from_bytes(package(constants::ODF_CHART, xml)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_and_excessive_depth() {
        let duplicate = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:x="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body><o:chart><c:chart c:class="c:bar" x:class="c:line"><c:plot-area/></c:chart></o:chart></o:body></o:document-content>"#;
        assert!(ChartDocument::from_bytes(package(constants::ODF_CHART, duplicate)).is_err());

        let nested = "<c:series>".repeat(129) + &"</c:series>".repeat(129);
        let deep = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:c="{CHART_NAMESPACE}"><o:body><o:chart><c:chart>{nested}</c:chart></o:chart></o:body></o:document-content>"#
        );
        assert!(ChartDocument::from_bytes(package(constants::ODF_CHART, &deep)).is_err());
    }
}
