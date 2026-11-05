//! Internal chart-reader state and namespace-preserving XML stream adapter.
//!
//! The adapter keeps the semantic parser focused on chart elements while
//! retaining self-contained `DrawingML` fragments and skipping unknown content
//! without changing the public chart model.

use crate::chart::Limits;
use crate::chart::limits::decoded_attribute_value_bytes;
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_ooxml_common::xml::attributes::{SeenNames, first_wins};
use litchi_ooxml_common::xml::{is_drawingml_chart_name, is_drawingml_name};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::{Config, NsReader};
use quick_xml::writer::Writer;
use std::io::BufRead;

pub(super) const IGNORED_NAMESPACE_ELEMENT: &str = "ignoredNamespaceElement";
pub(super) const INVALID_COLOR_MAPPING_ELEMENT: &str = "invalidColorMappingElement";

/// Namespace-aware streaming adapter for the chart model parser.
///
/// Core chart elements are exposed unchanged. `DrawingML` text and color-map choice
/// elements are also kept so their typed models can be decoded, while all other
/// namespaces are skipped as extension content. Rewriting the remaining `DrawingML`
/// container names prevents them from being mistaken for same-local-name chart
/// elements by the focused parsers below.
pub(super) struct ChartXmlReader<R: BufRead> {
    inner: NsReader<R>,
    limits: Limits,
    nodes: usize,
    depth: usize,
    skipped_depth: usize,
    saw_root: bool,
    closed_root: bool,
    series: usize,
    points_in_series: usize,
    total_points: usize,
    categories: usize,
    root_namespace_attributes: Vec<(Vec<u8>, Vec<u8>)>,
}

impl<R: BufRead> ChartXmlReader<R> {
    pub(super) fn from_reader(reader: R, limits: Limits) -> Self {
        Self {
            inner: NsReader::from_reader(reader),
            limits,
            nodes: 0,
            depth: 0,
            skipped_depth: 0,
            saw_root: false,
            closed_root: false,
            series: 0,
            points_in_series: 0,
            total_points: 0,
            categories: 0,
            root_namespace_attributes: Vec::new(),
        }
    }

    pub(super) fn config_mut(&mut self) -> &mut Config {
        self.inner.config_mut()
    }

    pub(super) fn decoder(&self) -> Decoder {
        self.inner.decoder()
    }

    pub(super) fn max_text_bytes(&self) -> usize {
        self.limits.max_text_bytes
    }

    pub(super) fn max_points(&self) -> usize {
        self.limits.max_points
    }

    /// Append decoded semantic text without allowing a caller-selected text
    /// field to grow past its bound.  XML text may arrive as several events
    /// (or as entity references), so checking only the final model would make
    /// the temporary `String` larger than the active policy first.
    pub(super) fn append_text(&self, target: &mut String, text: &str) -> Result<()> {
        let observed = target.len().checked_add(text.len()).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::ChartTextBytes,
                usize::MAX,
                self.limits.max_text_bytes,
            )
        })?;
        if observed > self.limits.max_text_bytes {
            return Err(Error::limit(
                crate::LimitKind::ChartTextBytes,
                observed,
                self.limits.max_text_bytes,
            ));
        }
        target.push_str(text);
        Ok(())
    }

    /// Admit one semantic series before its parser allocates the projection.
    pub(super) fn begin_series(&mut self) -> Result<()> {
        self.series = self.series.checked_add(1).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::ChartSeries,
                usize::MAX,
                self.limits.max_series,
            )
        })?;
        if self.series > self.limits.max_series {
            return Err(Error::limit(
                crate::LimitKind::ChartSeries,
                self.series,
                self.limits.max_series,
            ));
        }
        self.points_in_series = 0;
        // Category caches are retained once per series in the classic chart
        // model.  A repeated category dimension therefore starts a fresh
        // per-series budget, just as `validate_chart` does below.
        self.categories = 0;
        Ok(())
    }

    /// Admit one category point before its decoded string is retained.
    pub(super) fn admit_category_point(&mut self) -> Result<()> {
        self.categories = self.categories.checked_add(1).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::ChartCategories,
                usize::MAX,
                self.limits.max_categories,
            )
        })?;
        if self.categories > self.limits.max_categories {
            return Err(Error::limit(
                crate::LimitKind::ChartCategories,
                self.categories,
                self.limits.max_categories,
            ));
        }
        Ok(())
    }

    /// Admit one numeric point before its decoded value is retained.
    pub(super) fn admit_numeric_point(&mut self) -> Result<()> {
        self.points_in_series = self.points_in_series.checked_add(1).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::ChartPoints,
                usize::MAX,
                self.limits.max_points,
            )
        })?;
        if self.points_in_series > self.limits.max_points {
            return Err(Error::limit(
                crate::LimitKind::ChartPoints,
                self.points_in_series,
                self.limits.max_points,
            ));
        }
        self.total_points = self.total_points.checked_add(1).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::ChartPoints,
                usize::MAX,
                self.limits.max_total_points,
            )
        })?;
        if self.total_points > self.limits.max_total_points {
            return Err(Error::limit(
                crate::LimitKind::ChartPoints,
                self.total_points,
                self.limits.max_total_points,
            ));
        }
        Ok(())
    }

    pub(super) fn relationship_attribute_value(
        &self,
        element: &BytesStart<'_>,
        name: &[u8],
    ) -> Result<Option<String>> {
        const RELATIONSHIPS_NAMESPACE: &[u8] =
            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const STRICT_RELATIONSHIPS_NAMESPACE: &[u8] =
            b"http://purl.oclc.org/ooxml/officeDocument/relationships";

        let mut value = None;
        for attribute in element.checked_attributes() {
            let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
            if attribute.key.local_name().as_ref() != name {
                continue;
            }
            let (namespace, _) = self.inner.resolver().resolve_attribute(attribute.key);
            let is_relationship = matches!(
                namespace,
                ResolveResult::Bound(Namespace(value))
                    if value == RELATIONSHIPS_NAMESPACE
                        || value == STRICT_RELATIONSHIPS_NAMESPACE
            ) || matches!(namespace, ResolveResult::Unknown(prefix) if prefix.as_slice() == b"r");
            if !is_relationship {
                continue;
            }
            if value.is_some() {
                return Err(Error::Invalid(
                    "chart element contains duplicate relationship IDs".into(),
                ));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, self.decoder())
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .into_owned(),
            );
        }
        Ok(value)
    }

    pub(super) fn make_fragment_root_self_contained(
        &self,
        element: &BytesStart<'_>,
    ) -> BytesStart<'static> {
        let mut root = element.to_owned();
        // The first occurrence of each name counts and later duplicates are
        // skipped, at a cost of O(n log n) however the tag repeats its names.
        let mut existing_names = SeenNames::new();
        for attribute in first_wins(element).flatten() {
            existing_names.insert(attribute.key.into_inner());
        }
        for (name, value) in &self.root_namespace_attributes {
            if !existing_names.contains(&name.as_slice()) {
                root.push_attribute((name.as_slice(), value.as_slice()));
            }
        }
        root
    }

    pub(super) fn capture_empty_fragment(&self, element: &BytesStart<'_>) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Empty(
                self.make_fragment_root_self_contained(element),
            ))
            .map_err(|error| Error::Xml(error.to_string()))?;
        Ok(writer.into_inner())
    }

    pub(super) fn capture_fragment(
        &mut self,
        element: &BytesStart<'_>,
        description: &str,
    ) -> Result<Vec<u8>> {
        let mut writer = Writer::new(Vec::new());
        writer
            .write_event(Event::Start(
                self.make_fragment_root_self_contained(element),
            ))
            .map_err(|error| Error::Xml(error.to_string()))?;
        let fragment_depth = self.depth;
        let mut buffer = Vec::new();
        loop {
            let (_, event) = self
                .inner
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::Xml(error.to_string()))?;
            match event {
                Event::Start(ref element) => {
                    validate_attributes(element, self.limits, self.decoder())?;
                    let limits = self.limits;
                    count_node_values(&mut self.nodes, limits)?;
                    self.depth = self.depth.checked_add(1).ok_or_else(|| {
                        Error::limit(
                            crate::LimitKind::XmlDepth,
                            usize::MAX,
                            self.limits.max_xml_depth,
                        )
                    })?;
                    if self.depth > self.limits.max_xml_depth {
                        return Err(Error::limit(
                            crate::LimitKind::XmlDepth,
                            self.depth,
                            self.limits.max_xml_depth,
                        ));
                    }
                },
                Event::End(_) => {
                    self.depth = self.depth.checked_sub(1).ok_or_else(|| {
                        Error::Invalid(format!("{description} has an unmatched closing element"))
                    })?;
                },
                Event::Eof => {
                    return Err(Error::Invalid(format!("unterminated {description}")));
                },
                Event::Empty(ref element) => {
                    validate_attributes(element, self.limits, self.decoder())?;
                    let limits = self.limits;
                    count_node_values(&mut self.nodes, limits)?;
                    let depth = self.depth.checked_add(1).ok_or_else(|| {
                        Error::limit(
                            crate::LimitKind::XmlDepth,
                            usize::MAX,
                            self.limits.max_xml_depth,
                        )
                    })?;
                    if depth > self.limits.max_xml_depth {
                        return Err(Error::limit(
                            crate::LimitKind::XmlDepth,
                            depth,
                            self.limits.max_xml_depth,
                        ));
                    }
                },
                Event::DocType(_) => {
                    return Err(Error::Invalid(format!(
                        "{description} cannot contain a document type"
                    )));
                },
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
            let finished = matches!(event, Event::End(_)) && self.depth < fragment_depth;
            writer
                .write_event(event.into_owned())
                .map_err(|error| Error::Xml(error.to_string()))?;
            buffer.clear();
            if finished {
                break;
            }
        }
        Ok(writer.into_inner())
    }

    pub(super) fn read_event_into<'buffer>(
        &mut self,
        buffer: &'buffer mut Vec<u8>,
    ) -> Result<Event<'buffer>> {
        let limits = self.limits;
        let decoder = self.inner.decoder();
        let (namespace, event) = self
            .inner
            .read_resolved_event_into(buffer)
            .map_err(|error| Error::Xml(error.to_string()))?;

        match event {
            Event::Start(mut element) => {
                validate_attributes(&element, limits, decoder)?;
                count_node_values(&mut self.nodes, limits)?;
                let is_chart = is_chart_namespace(&namespace, &element);
                let is_drawing = is_drawing_namespace(&namespace, &element);

                if self.depth == 0 {
                    if self.saw_root
                        || !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace")
                    {
                        return Err(Error::Invalid(
                            "chart XML must have one DrawingML chartSpace root".to_string(),
                        ));
                    }
                    self.saw_root = true;
                    self.root_namespace_attributes.clear();
                    for attribute in element.checked_attributes() {
                        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                        let name = attribute.key.as_ref();
                        if name == b"xmlns" || name.starts_with(b"xmlns:") {
                            self.root_namespace_attributes
                                .push((name.to_vec(), attribute.value.into_owned()));
                        }
                    }
                    for (name, value) in [
                        (
                            b"xmlns:c".as_slice(),
                            b"http://schemas.openxmlformats.org/drawingml/2006/chart".as_slice(),
                        ),
                        (
                            b"xmlns:a".as_slice(),
                            b"http://schemas.openxmlformats.org/drawingml/2006/main".as_slice(),
                        ),
                        (
                            b"xmlns:r".as_slice(),
                            b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                                .as_slice(),
                        ),
                    ] {
                        if !self
                            .root_namespace_attributes
                            .iter()
                            .any(|(existing, _)| existing == name)
                        {
                            self.root_namespace_attributes
                                .push((name.to_vec(), value.to_vec()));
                        }
                    }
                }
                self.depth = self.depth.checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::XmlDepth,
                        usize::MAX,
                        self.limits.max_xml_depth,
                    )
                })?;
                if self.depth > self.limits.max_xml_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        self.depth,
                        self.limits.max_xml_depth,
                    ));
                }

                if self.skipped_depth > 0 {
                    self.skipped_depth = self.skipped_depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("chart XML nesting is too deep".to_string())
                    })?;
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_chart && is_drawing_color_map_choice(element.local_name().as_ref()) {
                    element.set_name(INVALID_COLOR_MAPPING_ELEMENT.as_bytes());
                } else if !is_chart && !is_drawing {
                    self.skipped_depth = self.skipped_depth.checked_add(1).ok_or_else(|| {
                        Error::Invalid("chart XML nesting is too deep".to_string())
                    })?;
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_drawing && !is_preserved_drawing_element(element.local_name().as_ref())
                {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                }
                Ok(Event::Start(element))
            },
            Event::Empty(mut element) => {
                validate_attributes(&element, limits, decoder)?;
                count_node_values(&mut self.nodes, limits)?;
                if self.depth == 0 {
                    return Err(Error::Invalid(
                        "chart XML must have one non-empty DrawingML chartSpace root".to_string(),
                    ));
                }
                let depth = self.depth.checked_add(1).ok_or_else(|| {
                    Error::limit(
                        crate::LimitKind::XmlDepth,
                        usize::MAX,
                        self.limits.max_xml_depth,
                    )
                })?;
                if depth > self.limits.max_xml_depth {
                    return Err(Error::limit(
                        crate::LimitKind::XmlDepth,
                        depth,
                        self.limits.max_xml_depth,
                    ));
                }
                let is_chart = is_chart_namespace(&namespace, &element);
                let is_drawing = is_drawing_namespace(&namespace, &element);
                if self.skipped_depth > 0 {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                } else if is_chart && is_drawing_color_map_choice(element.local_name().as_ref()) {
                    element.set_name(INVALID_COLOR_MAPPING_ELEMENT.as_bytes());
                } else if !is_chart
                    && (!is_drawing || !is_preserved_drawing_element(element.local_name().as_ref()))
                {
                    element.set_name(IGNORED_NAMESPACE_ELEMENT.as_bytes());
                }
                Ok(Event::Empty(element))
            },
            Event::End(element) => {
                if self.depth == 0 {
                    return Err(Error::Invalid(
                        "chart XML has an unmatched closing element".to_string(),
                    ));
                }
                self.depth -= 1;
                if self.skipped_depth > 0 {
                    self.skipped_depth -= 1;
                    return Ok(Event::End(BytesEnd::new(IGNORED_NAMESPACE_ELEMENT)));
                }

                let is_chart = is_drawingml_chart_name(
                    &namespace,
                    element.name(),
                    element.local_name().as_ref(),
                );
                let is_drawing =
                    is_drawingml_name(&namespace, element.name(), element.local_name().as_ref());
                if self.depth == 0 {
                    if !is_drawingml_chart_name(&namespace, element.name(), b"chartSpace") {
                        return Err(Error::Invalid(
                            "chart XML has an invalid root closing element".to_string(),
                        ));
                    }
                    self.closed_root = true;
                }
                if is_drawing && !is_preserved_drawing_element(element.local_name().as_ref()) {
                    return Ok(Event::End(BytesEnd::new(IGNORED_NAMESPACE_ELEMENT)));
                }
                if is_chart || is_drawing {
                    return Ok(Event::End(element));
                }
                Err(Error::Invalid(
                    "chart XML namespace state is inconsistent".to_string(),
                ))
            },
            _ if self.skipped_depth > 0 => Ok(Event::Comment(BytesText::new(""))),
            Event::Text(ref text) if self.depth == 0 => {
                if !text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(Error::Invalid(
                        "chart XML contains text outside its root".to_string(),
                    ));
                }
                Ok(event)
            },
            Event::CData(_) | Event::GeneralRef(_) if self.depth == 0 => Err(Error::Invalid(
                "chart XML contains data outside its root".to_string(),
            )),
            Event::Eof if !self.saw_root || !self.closed_root => Err(Error::Invalid(
                "chart XML has no complete chartSpace root".to_string(),
            )),
            Event::Eof => Ok(Event::Eof),
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => Ok(event),
        }
    }
}

fn count_node_values(nodes: &mut usize, limits: Limits) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        Error::limit(crate::LimitKind::XmlNodes, usize::MAX, limits.max_xml_nodes)
    })?;
    if *nodes > limits.max_xml_nodes {
        return Err(Error::limit(
            crate::LimitKind::XmlNodes,
            *nodes,
            limits.max_xml_nodes,
        ));
    }
    Ok(())
}

fn validate_attributes(element: &BytesStart<'_>, limits: Limits, decoder: Decoder) -> Result<()> {
    // The chart reader intentionally keeps the legacy first-wins behavior for
    // duplicate attributes.  The checked iterator rejects those duplicates,
    // so use the raw iterator here and only enforce resource budgets.
    let mut count = 0usize;
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        count = count.checked_add(1).ok_or_else(|| {
            Error::limit(
                crate::LimitKind::XmlAttributes,
                usize::MAX,
                limits.max_xml_attributes,
            )
        })?;
        if count > limits.max_xml_attributes {
            return Err(Error::limit(
                crate::LimitKind::XmlAttributes,
                count,
                limits.max_xml_attributes,
            ));
        }
        let value_length =
            decoded_attribute_value_bytes(attribute.value.as_ref(), decoder).map_err(Error::Xml)?;
        if value_length > limits.max_attribute_value_bytes {
            return Err(Error::limit(
                crate::LimitKind::AttributeValueBytes,
                value_length,
                limits.max_attribute_value_bytes,
            ));
        }
    }
    Ok(())
}

pub(super) fn is_chart_namespace(namespace: &ResolveResult<'_>, element: &BytesStart<'_>) -> bool {
    is_drawingml_chart_name(namespace, element.name(), element.local_name().as_ref())
}

pub(super) fn is_drawing_namespace(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
) -> bool {
    is_drawingml_name(namespace, element.name(), element.local_name().as_ref())
}

pub(super) fn is_preserved_drawing_element(local_name: &[u8]) -> bool {
    local_name == b"t" || is_drawing_color_map_choice(local_name)
}

pub(super) fn is_drawing_color_map_choice(local_name: &[u8]) -> bool {
    matches!(local_name, b"masterClrMapping" | b"overrideClrMapping")
}
