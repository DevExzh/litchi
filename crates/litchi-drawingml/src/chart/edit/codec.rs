//! Bounded XML ranges used by the chart editor.
//!
//! This is intentionally a byte-range editor rather than a second chart
//! serializer. Changed records are rebuilt only at the selected semantic
//! leaf/container; every other byte, including unknown extensions, prefixes,
//! attribute order, whitespace, and cached values, remains untouched.

use super::model::DataLabelFlag;
use super::validation;
use crate::chart::data::TitleText;
use crate::chart::reader;
use crate::chart::types::{AxisPosition, DisplayBlanks};
use crate::{Error, Result};
use litchi_core::xml::{escape_xml, unescape_xml};
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::io::Cursor;
use std::ops::Range;

const CHART_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/chart";
const DRAWING_TRANSITIONAL: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWING_STRICT: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Chart,
    Drawing,
}

#[derive(Debug, Clone)]
struct Attribute {
    name: Vec<u8>,
    value: Vec<u8>,
    value_range: Range<usize>,
}

#[derive(Debug, Clone)]
struct Node {
    qname: Vec<u8>,
    local: Vec<u8>,
    start: usize,
    open_end: usize,
    close_start: Option<usize>,
    end: usize,
    parent: Option<usize>,
}

impl Node {
    fn is_empty(&self) -> bool {
        self.close_start.is_none()
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Document {
    nodes: Vec<Node>,
    root: usize,
    namespaces: Vec<(Vec<u8>, Vec<u8>)>,
}

impl Document {
    pub(crate) fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > validation::MAX_XML_BYTES {
            return Err(validation::limit(
                "chart editor XML",
                validation::MAX_XML_BYTES,
            ));
        }

        let mut reader = Reader::from_reader(Cursor::new(xml));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack = Vec::new();
        let mut nodes = Vec::new();

        loop {
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(|error| Error::Xml(error.to_string()))?;
            let end = usize::try_from(reader.buffer_position())
                .map_err(|_| Error::Invalid("chart XML position exceeds usize".into()))?;
            match event {
                Event::Start(element) => {
                    if nodes.len() >= validation::MAX_XML_NODES {
                        return Err(validation::limit(
                            "chart editor XML nodes",
                            validation::MAX_XML_NODES,
                        ));
                    }
                    if stack.len() >= validation::MAX_XML_DEPTH {
                        return Err(validation::limit(
                            "chart editor XML nesting",
                            validation::MAX_XML_DEPTH,
                        ));
                    }
                    let start = tag_start(xml, end)?;
                    let index = nodes.len();
                    nodes.push(Node {
                        qname: element.name().as_ref().to_vec(),
                        local: element.local_name().as_ref().to_vec(),
                        start,
                        open_end: end,
                        close_start: None,
                        end: 0,
                        parent: stack.last().copied(),
                    });
                    stack.push(index);
                },
                Event::Empty(element) => {
                    if nodes.len() >= validation::MAX_XML_NODES {
                        return Err(validation::limit(
                            "chart editor XML nodes",
                            validation::MAX_XML_NODES,
                        ));
                    }
                    let start = tag_start(xml, end)?;
                    nodes.push(Node {
                        qname: element.name().as_ref().to_vec(),
                        local: element.local_name().as_ref().to_vec(),
                        start,
                        open_end: end,
                        close_start: None,
                        end,
                        parent: stack.last().copied(),
                    });
                },
                Event::End(element) => {
                    let Some(index) = stack.pop() else {
                        return Err(Error::Invalid(
                            "chart editor XML has an unmatched closing element".into(),
                        ));
                    };
                    if nodes[index].local != element.local_name().as_ref() {
                        return Err(Error::Invalid(
                            "chart editor XML closing element does not match its opener".into(),
                        ));
                    }
                    nodes[index].close_start = Some(tag_start(xml, end)?);
                    nodes[index].end = end;
                },
                Event::DocType(_) => {
                    return Err(Error::Invalid(
                        "chart editor XML cannot contain a document type".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
            buffer.clear();
        }

        if !stack.is_empty() {
            return Err(Error::Invalid("chart editor XML is not closed".into()));
        }
        let roots: Vec<_> = nodes
            .iter()
            .enumerate()
            .filter(|(_, node)| node.parent.is_none())
            .map(|(index, _)| index)
            .collect();
        if roots.len() != 1 || nodes[roots[0]].local != b"chartSpace" {
            return Err(Error::Invalid(
                "chart editor XML must have one chartSpace root".into(),
            ));
        }
        let root = roots[0];
        let namespaces = parse_namespace_attributes(xml, &nodes[root])?;
        Ok(Self {
            nodes,
            root,
            namespaces,
        })
    }

    pub(crate) fn root(&self) -> usize {
        self.root
    }

    fn node(&self, index: usize) -> &Node {
        &self.nodes[index]
    }

    fn children(&self, parent: usize, local: &[u8], kind: NamespaceKind) -> Vec<usize> {
        self.nodes
            .iter()
            .enumerate()
            .filter(|(_, node)| {
                node.parent == Some(parent)
                    && node.local == local
                    && self.matches_namespace(node, kind)
            })
            .map(|(index, _)| index)
            .collect()
    }

    fn child(&self, parent: usize, local: &[u8], kind: NamespaceKind) -> Result<Option<usize>> {
        let children = self.children(parent, local, kind);
        if children.len() > 1 {
            return Err(Error::Invalid(format!(
                "chart contains duplicate {} elements",
                String::from_utf8_lossy(local)
            )));
        }
        Ok(children.into_iter().next())
    }

    fn descendants(&self, parent: usize, local: &[u8], kind: NamespaceKind) -> Vec<usize> {
        self.nodes
            .iter()
            .enumerate()
            .filter(|(index, node)| {
                *index != parent
                    && node.local == local
                    && self.is_descendant(*index, parent)
                    && self.matches_namespace(node, kind)
            })
            .map(|(index, _)| index)
            .collect()
    }

    fn is_descendant(&self, mut child: usize, parent: usize) -> bool {
        while let Some(ancestor) = self.nodes[child].parent {
            if ancestor == parent {
                return true;
            }
            child = ancestor;
        }
        false
    }

    fn matches_namespace(&self, node: &Node, kind: NamespaceKind) -> bool {
        match kind {
            NamespaceKind::Chart => {
                self.namespace_uri(&node.qname) == Some(CHART_TRANSITIONAL)
                    || self.namespace_uri(&node.qname) == Some(CHART_STRICT)
            },
            NamespaceKind::Drawing => {
                self.namespace_uri(&node.qname) == Some(DRAWING_TRANSITIONAL)
                    || self.namespace_uri(&node.qname) == Some(DRAWING_STRICT)
            },
        }
    }

    fn namespace_uri(&self, qname: &[u8]) -> Option<&[u8]> {
        let prefix = qname
            .iter()
            .position(|byte| *byte == b':')
            .map(|position| &qname[..position])
            .unwrap_or_default();
        self.namespaces
            .iter()
            .find(|(name, _)| name.as_slice() == prefix)
            .map(|(_, value)| value.as_slice())
    }

    fn prefix(&self, index: usize) -> Vec<u8> {
        self.node(index)
            .qname
            .iter()
            .position(|byte| *byte == b':')
            .map(|position| self.node(index).qname[..position].to_vec())
            .unwrap_or_default()
    }

    fn namespace_prefix(&self, kind: NamespaceKind) -> Option<Vec<u8>> {
        self.namespaces
            .iter()
            .find(|(_, uri)| match kind {
                NamespaceKind::Chart => {
                    uri.as_slice() == CHART_TRANSITIONAL || uri.as_slice() == CHART_STRICT
                },
                NamespaceKind::Drawing => {
                    uri.as_slice() == DRAWING_TRANSITIONAL || uri.as_slice() == DRAWING_STRICT
                },
            })
            .map(|(prefix, _)| prefix.clone())
    }

    fn chart(&self) -> Result<usize> {
        self.child(self.root, b"chart", NamespaceKind::Chart)?
            .ok_or_else(|| Error::Invalid("chart XML has no chart element".into()))
    }

    fn plot_area(&self) -> Result<usize> {
        self.child(self.chart()?, b"plotArea", NamespaceKind::Chart)?
            .ok_or_else(|| Error::Invalid("chart XML has no plotArea element".into()))
    }

    fn series(&self, xml: &[u8], index: u32) -> Result<usize> {
        let plot_area = self.plot_area()?;
        let mut matches = Vec::new();
        for (candidate, node) in self.nodes.iter().enumerate() {
            if node.local != b"ser"
                || !self.matches_namespace(node, NamespaceKind::Chart)
                || !self.is_descendant(candidate, plot_area)
            {
                continue;
            }
            let Some(idx) = self.child(candidate, b"idx", NamespaceKind::Chart)? else {
                continue;
            };
            if self
                .attribute_from(xml, idx, b"val")?
                .and_then(|value| parse_u32(&value).ok())
                == Some(index)
            {
                matches.push(candidate);
            }
        }
        match matches.as_slice() {
            [] => Err(Error::Invalid(format!(
                "chart series index {index} was not found"
            ))),
            [series] => Ok(*series),
            _ => Err(Error::Invalid(format!(
                "chart contains duplicate series index {index}"
            ))),
        }
    }

    fn axis(&self, xml: &[u8], axis_id: u32) -> Result<usize> {
        let plot_area = self.plot_area()?;
        let mut matches = Vec::new();
        for (candidate, node) in self.nodes.iter().enumerate() {
            if !matches!(
                node.local.as_slice(),
                b"catAx" | b"dateAx" | b"serAx" | b"valAx"
            ) || !self.matches_namespace(node, NamespaceKind::Chart)
                || !self.is_descendant(candidate, plot_area)
            {
                continue;
            }
            let Some(id) = self.child(candidate, b"axId", NamespaceKind::Chart)? else {
                continue;
            };
            if self
                .attribute_from(xml, id, b"val")?
                .and_then(|value| parse_u32(&value).ok())
                == Some(axis_id)
            {
                matches.push(candidate);
            }
        }
        match matches.as_slice() {
            [] => Err(Error::Invalid(format!(
                "chart axis ID {axis_id} was not found"
            ))),
            [axis] => Ok(*axis),
            _ => Err(Error::Invalid(format!(
                "chart contains duplicate axis ID {axis_id}"
            ))),
        }
    }

    fn attribute_from(&self, xml: &[u8], index: usize, name: &[u8]) -> Result<Option<Vec<u8>>> {
        let node = self.node(index);
        Ok(
            parse_attributes_with_base(&xml[node.start..node.open_end], node.start)?
                .into_iter()
                .find(|attribute| attribute.name == name)
                .map(|attribute| decode_attribute(&attribute.value)),
        )
    }

    fn find_attribute(&self, xml: &[u8], index: usize, name: &[u8]) -> Result<Option<Attribute>> {
        let node = self.node(index);
        Ok(
            parse_attributes_with_base(&xml[node.start..node.open_end], node.start)?
                .into_iter()
                .find(|attribute| attribute.name == name),
        )
    }
}

pub(crate) fn read(xml: &[u8]) -> Result<crate::chart::Chart> {
    Document::parse(xml)?;
    reader::read(Cursor::new(xml))
}

pub(crate) fn set_style(xml: &[u8], style: Option<u32>) -> Result<Vec<u8>> {
    validation::validate_style(style)?;
    let document = Document::parse(xml)?;
    let node = document.child(document.root(), b"style", NamespaceKind::Chart)?;
    match (node, style) {
        (Some(node), Some(style)) => {
            replace_attribute_if_changed(xml, &document, node, b"val", &style.to_string())
        },
        (Some(node), None) => remove_node(xml, &document, node),
        (None, Some(style)) => {
            let prefix = document.prefix(document.root());
            let fragment = element(&prefix, "style", &format!(r#" val="{style}""#), "");
            insert_before_child(xml, &document, document.root(), b"chart", &fragment)
        },
        (None, None) => Ok(xml.to_vec()),
    }
}

pub(crate) fn set_language(xml: &[u8], language: Option<&str>) -> Result<Vec<u8>> {
    if let Some(language) = language {
        validation::validate_text(language, "language text")?;
    }
    let document = Document::parse(xml)?;
    let node = document.child(document.root(), b"lang", NamespaceKind::Chart)?;
    match (node, language) {
        (Some(node), Some(language)) => {
            replace_attribute_if_changed(xml, &document, node, b"val", language)
        },
        (Some(node), None) => remove_node(xml, &document, node),
        (None, Some(language)) => {
            let prefix = document.prefix(document.root());
            let fragment = element(
                &prefix,
                "lang",
                &format!(r#" val="{}""#, escape_xml(language)),
                "",
            );
            insert_before_child(xml, &document, document.root(), b"chart", &fragment)
        },
        (None, None) => Ok(xml.to_vec()),
    }
}

pub(crate) fn set_display_blanks(xml: &[u8], mode: DisplayBlanks) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let chart = document.chart()?;
    let node = document.child(chart, b"dispBlanksAs", NamespaceKind::Chart)?;
    let value = mode.xml_value();
    match node {
        Some(node) => replace_attribute_if_changed(xml, &document, node, b"val", value),
        None => {
            let prefix = document.prefix(chart);
            let fragment = element(&prefix, "dispBlanksAs", &format!(r#" val="{value}""#), "");
            append_to_parent(xml, &document, chart, &fragment)
        },
    }
}

pub(crate) fn set_chart_title(xml: &[u8], title: Option<&TitleText>) -> Result<Vec<u8>> {
    let mut source = xml.to_vec();
    let mut document = Document::parse(&source)?;
    let chart = document.chart()?;
    let title_node = document.child(chart, b"title", NamespaceKind::Chart)?;
    match (title_node, title) {
        (Some(node), None) => remove_node(&source, &document, node),
        (Some(node), Some(title)) => set_title_in_node(&source, &document, node, title),
        (None, Some(title)) => {
            if matches!(title, TitleText::Literal(_))
                && document.namespace_prefix(NamespaceKind::Drawing).is_none()
            {
                source = ensure_namespace(&source, &document, b"a", DRAWING_TRANSITIONAL)?;
                document = Document::parse(&source)?;
            }
            let prefix = document.prefix(chart);
            let drawing = document
                .namespace_prefix(NamespaceKind::Drawing)
                .unwrap_or_else(|| b"a".to_vec());
            let fragment = title_fragment(&prefix, &drawing, title);
            insert_before_child(&source, &document, chart, b"plotArea", &fragment)
        },
        (None, None) => Ok(source),
    }
}

pub(crate) fn set_series_title(
    xml: &[u8],
    index: u32,
    title: Option<&TitleText>,
) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let series = document.series(xml, index)?;
    let tx = document.child(series, b"tx", NamespaceKind::Chart)?;
    match (tx, title) {
        (Some(tx), None) => remove_node(xml, &document, tx),
        (Some(tx), Some(title)) => set_title_in_node(xml, &document, tx, title),
        (None, Some(title)) => {
            let prefix = document.prefix(series);
            let fragment = match title {
                TitleText::Literal(value) => element(
                    &prefix,
                    "tx",
                    "",
                    &element(&prefix, "v", "", &escape_xml(&value.text)),
                ),
                TitleText::Reference(value) => element(
                    &prefix,
                    "tx",
                    "",
                    &element(&prefix, "f", "", &escape_xml(&value.formula)),
                ),
            };
            insert_after_child(xml, &document, series, b"order", &fragment)
        },
        (None, None) => Ok(xml.to_vec()),
    }
}

pub(crate) fn set_series_data_label_flag(
    xml: &[u8],
    index: u32,
    flag: DataLabelFlag,
    value: bool,
) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let series = document.series(xml, index)?;
    let labels = document.child(series, b"dLbls", NamespaceKind::Chart)?;
    let name = flag.element().as_bytes();
    match labels {
        Some(labels) => match document.child(labels, name, NamespaceKind::Chart)? {
            Some(node) => {
                let lexical = if value { "1" } else { "0" };
                replace_attribute_if_changed(xml, &document, node, b"val", lexical)
            },
            None => {
                let prefix = document.prefix(labels);
                let fragment = element(
                    &prefix,
                    flag.element(),
                    &format!(r#" val="{}""#, if value { 1 } else { 0 }),
                    "",
                );
                append_to_parent(xml, &document, labels, &fragment)
            },
        },
        None => {
            let prefix = document.prefix(series);
            let child = element(
                &prefix,
                flag.element(),
                &format!(r#" val="{}""#, if value { 1 } else { 0 }),
                "",
            );
            let labels = element(&prefix, "dLbls", "", &child);
            append_to_parent(xml, &document, series, &labels)
        },
    }
}

pub(crate) fn clear_series_data_label_flag(
    xml: &[u8],
    index: u32,
    flag: DataLabelFlag,
) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let series = document.series(xml, index)?;
    let Some(labels) = document.child(series, b"dLbls", NamespaceKind::Chart)? else {
        return Ok(xml.to_vec());
    };
    let Some(node) = document.child(labels, flag.element().as_bytes(), NamespaceKind::Chart)?
    else {
        return Ok(xml.to_vec());
    };
    remove_node(xml, &document, node)
}

pub(crate) fn set_series_data_label_separator(
    xml: &[u8],
    index: u32,
    separator: Option<&str>,
) -> Result<Vec<u8>> {
    if let Some(separator) = separator {
        validation::validate_text(separator, "data-label separator")?;
    }
    let document = Document::parse(xml)?;
    let series = document.series(xml, index)?;
    let labels = document.child(series, b"dLbls", NamespaceKind::Chart)?;
    match (labels, separator) {
        (None, None) => Ok(xml.to_vec()),
        (Some(labels), None) => match document.child(labels, b"separator", NamespaceKind::Chart)? {
            Some(node) => remove_node(xml, &document, node),
            None => Ok(xml.to_vec()),
        },
        (Some(labels), Some(separator)) => {
            let fragment = escape_xml(separator);
            match document.child(labels, b"separator", NamespaceKind::Chart)? {
                Some(node) => replace_text(xml, &document, node, &fragment),
                None => {
                    let prefix = document.prefix(labels);
                    let element = element(&prefix, "separator", "", &fragment);
                    append_to_parent(xml, &document, labels, &element)
                },
            }
        },
        (None, Some(separator)) => {
            let prefix = document.prefix(series);
            let separator = element(&prefix, "separator", "", &escape_xml(separator));
            let labels = element(&prefix, "dLbls", "", &separator);
            append_to_parent(xml, &document, series, &labels)
        },
    }
}

pub(crate) fn set_axis_position(
    xml: &[u8],
    axis_id: u32,
    position: AxisPosition,
) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let axis = document.axis(xml, axis_id)?;
    let value = position.xml_value();
    match document.child(axis, b"axPos", NamespaceKind::Chart)? {
        Some(node) => replace_attribute_if_changed(xml, &document, node, b"val", value),
        None => {
            let prefix = document.prefix(axis);
            let fragment = element(&prefix, "axPos", &format!(r#" val="{value}""#), "");
            insert_after_child(xml, &document, axis, b"scaling", &fragment)
        },
    }
}

pub(crate) fn set_axis_range(
    xml: &[u8],
    axis_id: u32,
    min: Option<f64>,
    max: Option<f64>,
) -> Result<Vec<u8>> {
    validation::validate_axis_range(min, max)?;
    let document = Document::parse(xml)?;
    let axis = document.axis(xml, axis_id)?;
    document
        .child(axis, b"scaling", NamespaceKind::Chart)?
        .ok_or_else(|| Error::Invalid("chart axis has no scaling element".into()))?;

    let mut source = xml.to_vec();
    if let Some(max) = max {
        source = set_scaling_value(&source, axis_id, b"max", max)?;
    } else {
        source = remove_scaling_value(&source, axis_id, b"max")?;
    }
    if let Some(min) = min {
        source = set_scaling_value(&source, axis_id, b"min", min)?;
    } else {
        source = remove_scaling_value(&source, axis_id, b"min")?;
    }
    Ok(source)
}

fn set_scaling_value(xml: &[u8], axis_id: u32, name: &[u8], value: f64) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let axis = document.axis(xml, axis_id)?;
    let scaling = document
        .child(axis, b"scaling", NamespaceKind::Chart)?
        .ok_or_else(|| Error::Invalid("chart axis has no scaling element".into()))?;
    if let Some(node) = document.child(scaling, name, NamespaceKind::Chart)? {
        if document
            .attribute_from(xml, node, b"val")?
            .and_then(|value| parse_f64(&value).ok())
            == Some(value)
        {
            return Ok(xml.to_vec());
        }
        return replace_attribute_if_changed(xml, &document, node, b"val", &value.to_string());
    }
    let prefix = document.prefix(scaling);
    let local =
        std::str::from_utf8(name).map_err(|_| Error::Invalid("invalid axis field".into()))?;
    let fragment = element(&prefix, local, &format!(r#" val="{value}""#), "");
    if name == b"max"
        && let Some(min) = document.child(scaling, b"min", NamespaceKind::Chart)?
    {
        return insert_at(xml, document.node(min).start, fragment.as_bytes());
    }
    append_to_parent(xml, &document, scaling, &fragment)
}

fn remove_scaling_value(xml: &[u8], axis_id: u32, name: &[u8]) -> Result<Vec<u8>> {
    let document = Document::parse(xml)?;
    let axis = document.axis(xml, axis_id)?;
    let scaling = document
        .child(axis, b"scaling", NamespaceKind::Chart)?
        .ok_or_else(|| Error::Invalid("chart axis has no scaling element".into()))?;
    match document.child(scaling, name, NamespaceKind::Chart)? {
        Some(node) => remove_node(xml, &document, node),
        None => Ok(xml.to_vec()),
    }
}

fn set_title_in_node(
    xml: &[u8],
    document: &Document,
    node: usize,
    title: &TitleText,
) -> Result<Vec<u8>> {
    let tx = if document.node(node).local == b"tx" {
        node
    } else {
        document
            .child(node, b"tx", NamespaceKind::Chart)?
            .ok_or_else(|| Error::Invalid("chart title has no text container".into()))?
    };
    match title {
        TitleText::Literal(value) => {
            validation::validate_text(&value.text, "title text")?;
            if let Some(value_node) = document.child(tx, b"v", NamespaceKind::Chart)? {
                return replace_text(xml, document, value_node, &escape_xml(&value.text));
            }
            if let Some(rich) = document.child(tx, b"rich", NamespaceKind::Chart)? {
                let text = document
                    .descendants(rich, b"t", NamespaceKind::Drawing)
                    .into_iter()
                    .next()
                    .ok_or_else(|| {
                        Error::Invalid("chart title rich text has no text run".into())
                    })?;
                return replace_text(xml, document, text, &escape_xml(&value.text));
            }
            if document
                .child(tx, b"strRef", NamespaceKind::Chart)?
                .is_some()
            {
                return Err(Error::Invalid(
                    "changing a reference title to literal text would discard cached data".into(),
                ));
            }
            append_to_parent(
                xml,
                document,
                tx,
                &element(&document.prefix(tx), "v", "", &escape_xml(&value.text)),
            )
        },
        TitleText::Reference(value) => {
            validation::validate_text(&value.formula, "title formula")?;
            if let Some(reference) = document.child(tx, b"strRef", NamespaceKind::Chart)? {
                if let Some(formula) = document.child(reference, b"f", NamespaceKind::Chart)? {
                    return replace_text(xml, document, formula, &escape_xml(&value.formula));
                }
                return Err(Error::Invalid(
                    "chart reference title has no formula".into(),
                ));
            }
            if let Some(formula) = document.child(tx, b"f", NamespaceKind::Chart)? {
                return replace_text(xml, document, formula, &escape_xml(&value.formula));
            }
            if document.child(tx, b"rich", NamespaceKind::Chart)?.is_some()
                || document.child(tx, b"v", NamespaceKind::Chart)?.is_some()
            {
                return Err(Error::Invalid(
                    "changing a literal title to a reference would discard rich text".into(),
                ));
            }
            append_to_parent(
                xml,
                document,
                tx,
                &element(&document.prefix(tx), "f", "", &escape_xml(&value.formula)),
            )
        },
    }
}

fn title_fragment(prefix: &[u8], drawing_prefix: &[u8], title: &TitleText) -> String {
    let title_name = qname(prefix, "title");
    let tx_name = qname(prefix, "tx");
    let overlay_name = qname(prefix, "overlay");
    match title {
        TitleText::Literal(value) => {
            let rich_name = qname(prefix, "rich");
            let body = qname(drawing_prefix, "bodyPr");
            let list = qname(drawing_prefix, "lstStyle");
            let paragraph = qname(drawing_prefix, "p");
            let run = qname(drawing_prefix, "r");
            let text = qname(drawing_prefix, "t");
            format!(
                "<{title_name}><{tx_name}><{rich_name}><{body}/><{list}/><{paragraph}><{run}><{text}>{}</{text}></{run}></{paragraph}></{rich_name}></{tx_name}><{overlay_name} val=\"0\"/></{title_name}>",
                escape_xml(&value.text)
            )
        },
        TitleText::Reference(value) => {
            let reference = qname(prefix, "strRef");
            let formula = qname(prefix, "f");
            format!(
                "<{title_name}><{tx_name}><{reference}><{formula}>{}</{formula}></{reference}></{tx_name}><{overlay_name} val=\"0\"/></{title_name}>",
                escape_xml(&value.formula)
            )
        },
    }
}

fn qname(prefix: &[u8], local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{}:{local}", String::from_utf8_lossy(prefix))
    }
}

fn element(prefix: &[u8], local: &str, attributes: &str, content: &str) -> String {
    let name = qname(prefix, local);
    if content.is_empty() {
        format!("<{name}{attributes}/>")
    } else {
        format!("<{name}{attributes}>{content}</{name}>")
    }
}

fn tag_start(xml: &[u8], end: usize) -> Result<usize> {
    let mut index = end.min(xml.len());
    while index > 0 {
        index -= 1;
        if xml[index] == b'<' {
            return Ok(index);
        }
    }
    Err(Error::Invalid(
        "chart XML event has no opening delimiter".into(),
    ))
}

fn parse_namespace_attributes(xml: &[u8], root: &Node) -> Result<Vec<(Vec<u8>, Vec<u8>)>> {
    let mut namespaces = Vec::new();
    for attribute in parse_attributes(&xml[root.start..root.open_end])? {
        if attribute.name == b"xmlns" {
            namespaces.push((Vec::new(), decode_attribute(&attribute.value)));
        } else if let Some(prefix) = attribute.name.strip_prefix(b"xmlns:") {
            namespaces.push((prefix.to_vec(), decode_attribute(&attribute.value)));
        }
    }
    Ok(namespaces)
}

fn parse_attributes(raw: &[u8]) -> Result<Vec<Attribute>> {
    parse_attributes_with_base(raw, 0)
}

fn parse_attributes_with_base(raw: &[u8], base: usize) -> Result<Vec<Attribute>> {
    let mut attributes = Vec::new();
    let mut index = 1usize;
    while index < raw.len()
        && !raw[index].is_ascii_whitespace()
        && raw[index] != b'>'
        && raw[index] != b'/'
    {
        index += 1;
    }
    while index < raw.len() {
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= raw.len()
            || raw[index] == b'>'
            || (raw[index] == b'/' && raw.get(index + 1) == Some(&b'>'))
        {
            break;
        }
        let name_start = index;
        while index < raw.len()
            && !raw[index].is_ascii_whitespace()
            && !matches!(raw[index], b'=' | b'/' | b'>')
        {
            index += 1;
        }
        if name_start == index {
            return Err(Error::Invalid(
                "chart XML contains an invalid attribute name".into(),
            ));
        }
        let name = raw[name_start..index].to_vec();
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        if raw.get(index) != Some(&b'=') {
            return Err(Error::Invalid(format!(
                "chart XML attribute has no value at {index} after {:?} in {:?}",
                String::from_utf8_lossy(&raw[name_start..index]),
                String::from_utf8_lossy(raw)
            )));
        }
        index += 1;
        while index < raw.len() && raw[index].is_ascii_whitespace() {
            index += 1;
        }
        let quote = *raw
            .get(index)
            .ok_or_else(|| Error::Invalid("chart XML attribute is truncated".into()))?;
        if quote != b'"' && quote != b'\'' {
            return Err(Error::Invalid("chart XML attribute is not quoted".into()));
        }
        index += 1;
        let value_start = index;
        while index < raw.len() && raw[index] != quote {
            index += 1;
        }
        let value_end = index;
        if index >= raw.len() {
            return Err(Error::Invalid("chart XML attribute is unterminated".into()));
        }
        attributes.push(Attribute {
            name,
            value: raw[value_start..value_end].to_vec(),
            value_range: (base + value_start)..(base + value_end),
        });
        index += 1;
    }
    Ok(attributes)
}

fn decode_attribute(value: &[u8]) -> Vec<u8> {
    std::str::from_utf8(value)
        .map(|value| unescape_xml(value).into_bytes())
        .unwrap_or_else(|_| value.to_vec())
}

fn parse_u32(value: &[u8]) -> Result<u32> {
    std::str::from_utf8(value)
        .map_err(|_| Error::Invalid("chart numeric attribute is not UTF-8".into()))?
        .parse()
        .map_err(|_| Error::Invalid("chart numeric attribute is invalid".into()))
}

fn parse_f64(value: &[u8]) -> Result<f64> {
    let value = std::str::from_utf8(value)
        .map_err(|_| Error::Invalid("chart numeric attribute is not UTF-8".into()))?;
    let value = value
        .parse::<f64>()
        .map_err(|_| Error::Invalid("chart numeric attribute is invalid".into()))?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(Error::Invalid(
            "chart numeric attribute is not finite".into(),
        ))
    }
}

fn replace_attribute_if_changed(
    xml: &[u8],
    document: &Document,
    node: usize,
    name: &[u8],
    value: &str,
) -> Result<Vec<u8>> {
    let Some(attribute) = document.find_attribute(xml, node, name)? else {
        let open_end = document.node(node).open_end;
        let insertion = open_end
            - if xml.get(open_end.saturating_sub(2)) == Some(&b'/') {
                2
            } else {
                1
            };
        let name = String::from_utf8_lossy(name);
        return insert_at(
            xml,
            insertion,
            format!(r#" {name}="{}""#, escape_xml(value)).as_bytes(),
        );
    };
    if decode_attribute(&attribute.value) == value.as_bytes() {
        return Ok(xml.to_vec());
    }
    replace_range(xml, attribute.value_range, escape_xml(value).as_bytes())
}

fn replace_text(xml: &[u8], document: &Document, node: usize, value: &str) -> Result<Vec<u8>> {
    let target = document.node(node);
    if target.is_empty() {
        let prefix = target
            .qname
            .iter()
            .position(|byte| *byte == b':')
            .map(|position| target.qname[..position].to_vec())
            .unwrap_or_default();
        let local = std::str::from_utf8(&target.local)
            .map_err(|_| Error::Invalid("chart text element name is invalid".into()))?;
        let fragment = element(&prefix, local, "", value);
        return replace_range(xml, target.start..target.end, fragment.as_bytes());
    }
    let content_start = target.open_end;
    let content_end = target
        .close_start
        .ok_or_else(|| Error::Invalid("chart text element is not closed".into()))?;
    if xml[content_start..content_end].contains(&b'<') {
        return Err(Error::Invalid(
            "chart text edit would discard opaque child XML".into(),
        ));
    }
    if &xml[content_start..content_end] == value.as_bytes() {
        return Ok(xml.to_vec());
    }
    replace_range(xml, content_start..content_end, value.as_bytes())
}

fn remove_node(xml: &[u8], document: &Document, node: usize) -> Result<Vec<u8>> {
    replace_range(xml, document.node(node).start..document.node(node).end, &[])
}

fn append_to_parent(
    xml: &[u8],
    document: &Document,
    parent: usize,
    fragment: &str,
) -> Result<Vec<u8>> {
    let node = document.node(parent);
    if let Some(close_start) = node.close_start {
        return insert_at(xml, close_start, fragment.as_bytes());
    }
    if !node.is_empty() {
        return Err(Error::Invalid("chart parent element is not closed".into()));
    }
    let raw = &xml[node.start..node.open_end];
    let open = raw
        .strip_suffix(b"/>")
        .ok_or_else(|| Error::Invalid("chart empty element has no close marker".into()))?;
    let mut replacement = open.to_vec();
    replacement.push(b'>');
    replacement.extend_from_slice(fragment.as_bytes());
    replacement.extend_from_slice(b"</");
    replacement.extend_from_slice(&node.qname);
    replacement.push(b'>');
    replace_range(xml, node.start..node.end, &replacement)
}

fn insert_before_child(
    xml: &[u8],
    document: &Document,
    parent: usize,
    child: &[u8],
    fragment: &str,
) -> Result<Vec<u8>> {
    if let Some(child) = document.child(parent, child, NamespaceKind::Chart)? {
        return insert_at(xml, document.node(child).start, fragment.as_bytes());
    }
    append_to_parent(xml, document, parent, fragment)
}

fn insert_after_child(
    xml: &[u8],
    document: &Document,
    parent: usize,
    child: &[u8],
    fragment: &str,
) -> Result<Vec<u8>> {
    if let Some(child) = document.child(parent, child, NamespaceKind::Chart)? {
        return insert_at(xml, document.node(child).end, fragment.as_bytes());
    }
    append_to_parent(xml, document, parent, fragment)
}

fn ensure_namespace(xml: &[u8], document: &Document, prefix: &[u8], uri: &[u8]) -> Result<Vec<u8>> {
    if document.namespace_uri(prefix) == Some(uri) {
        return Ok(xml.to_vec());
    }
    let root = document.node(document.root());
    let insertion = root.open_end
        - if xml.get(root.open_end.saturating_sub(2)) == Some(&b'/') {
            2
        } else {
            1
        };
    let declaration = format!(
        r#" xmlns:{}="{}""#,
        String::from_utf8_lossy(prefix),
        String::from_utf8_lossy(uri)
    );
    insert_at(xml, insertion, declaration.as_bytes())
}

fn insert_at(xml: &[u8], position: usize, bytes: &[u8]) -> Result<Vec<u8>> {
    if position > xml.len() {
        return Err(Error::Invalid(
            "chart XML insertion position is out of bounds".into(),
        ));
    }
    let output_len = xml.len().saturating_add(bytes.len());
    if output_len > validation::MAX_XML_BYTES {
        return Err(validation::limit(
            "chart editor XML",
            validation::MAX_XML_BYTES,
        ));
    }
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(bytes);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
}

fn replace_range(xml: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > xml.len() {
        return Err(Error::Invalid(
            "chart XML replacement range is invalid".into(),
        ));
    }
    let output_len = xml
        .len()
        .saturating_sub(range.end - range.start)
        .saturating_add(replacement.len());
    if output_len > validation::MAX_XML_BYTES {
        return Err(validation::limit(
            "chart editor XML",
            validation::MAX_XML_BYTES,
        ));
    }
    let mut output = Vec::with_capacity(output_len);
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}
