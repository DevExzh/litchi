//! Typed, inert ODF drawing layer metadata.

use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};
use std::io::BufRead;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const SVG_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_SETS: usize = 65_536;
const MAX_LAYERS: usize = 262_144;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// The XML part containing a layer set.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfDrawingLayerPart {
    Content,
    Styles,
    FlatDocument,
}

/// The schema-defined container of a layer set.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfDrawingLayerScope {
    MasterStyles,
    Page { name: Option<String> },
}

/// The exact `draw:display` mode. An omitted value defaults to `always`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfDrawingLayerDisplay {
    Always,
    Screen,
    Printer,
    None,
}

impl OdfDrawingLayerDisplay {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "always" => Ok(Self::Always),
            "screen" => Ok(Self::Screen),
            "printer" => Ok(Self::Printer),
            "none" => Ok(Self::None),
            _ => invalid(format!("unsupported draw:display value '{value}'")),
        }
    }

    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Always => "always",
            Self::Screen => "screen",
            Self::Printer => "printer",
            Self::None => "none",
        }
    }
}

/// One ordered `draw:layer` declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDrawingLayer {
    pub name: String,
    pub protected: Option<bool>,
    pub display: Option<OdfDrawingLayerDisplay>,
    pub title: Option<String>,
    pub description: Option<String>,
}

impl OdfDrawingLayer {
    pub fn effective_protected(&self) -> bool {
        self.protected.unwrap_or(false)
    }

    pub fn effective_display(&self) -> OdfDrawingLayerDisplay {
        self.display.unwrap_or(OdfDrawingLayerDisplay::Always)
    }

    fn validate(&self) -> Result<()> {
        validate_text(&self.name, "draw:name", false)?;
        if let Some(title) = &self.title {
            validate_text(title, "svg:title", true)?;
        }
        if let Some(description) = &self.description {
            validate_text(description, "svg:desc", true)?;
        }
        Ok(())
    }
}

/// One global or page-local `draw:layer-set`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfDrawingLayerSet {
    pub part: OdfDrawingLayerPart,
    pub scope: OdfDrawingLayerScope,
    pub layers: Vec<OdfDrawingLayer>,
}

impl OdfDrawingLayerSet {
    pub fn get(&self, name: &str) -> Option<&OdfDrawingLayer> {
        self.layers.iter().find(|layer| layer.name == name)
    }

    pub fn validate(&self) -> Result<()> {
        if self.layers.len() > MAX_LAYERS {
            return invalid(format!("drawing layer set exceeds {MAX_LAYERS} layers"));
        }
        if let OdfDrawingLayerScope::Page { name: Some(name) } = &self.scope {
            validate_text(name, "draw:page draw:name", false)?;
        }
        let mut names = HashSet::with_capacity(self.layers.len());
        for layer in &self.layers {
            layer.validate()?;
            if !names.insert(layer.name.as_str()) {
                return invalid(format!("duplicate drawing layer name '{}'", layer.name));
            }
        }
        Ok(())
    }

    /// Serializes the layer set itself; its package part and scope remain metadata.
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192 + self.layers.len() * 96);
        output.push_str(r#"<draw:layer-set xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">"#);
        for layer in &self.layers {
            output.push_str("<draw:layer draw:name=\"");
            escape_attribute(&mut output, &layer.name);
            output.push('"');
            if let Some(protected) = layer.protected {
                output.push_str(" draw:protected=\"");
                output.push_str(if protected { "true" } else { "false" });
                output.push('"');
            }
            if let Some(display) = layer.display {
                output.push_str(" draw:display=\"");
                output.push_str(display.as_str());
                output.push('"');
            }
            if layer.title.is_none() && layer.description.is_none() {
                output.push_str("/>");
                continue;
            }
            output.push('>');
            if let Some(title) = &layer.title {
                output.push_str("<svg:title>");
                escape_text(&mut output, title);
                output.push_str("</svg:title>");
            }
            if let Some(description) = &layer.description {
                output.push_str("<svg:desc>");
                escape_text(&mut output, description);
                output.push_str("</svg:desc>");
            }
            output.push_str("</draw:layer>");
        }
        output.push_str("</draw:layer-set>");
        Ok(output)
    }
}

/// Ordered layer sets found across an ODF package.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfDrawingLayerSets {
    pub sets: Vec<OdfDrawingLayerSet>,
}

impl OdfDrawingLayerSets {
    pub fn validate(&self) -> Result<()> {
        if self.sets.len() > MAX_SETS {
            return invalid(format!("document exceeds {MAX_SETS} drawing layer sets"));
        }
        let mut layer_count = 0usize;
        let mut aggregate = 0usize;
        for set in &self.sets {
            set.validate()?;
            layer_count = layer_count
                .checked_add(set.layers.len())
                .ok_or_else(|| make_error("drawing layer count overflow"))?;
            if layer_count > MAX_LAYERS {
                return invalid(format!("document exceeds {MAX_LAYERS} drawing layers"));
            }
            if let OdfDrawingLayerScope::Page { name: Some(name) } = &set.scope {
                aggregate = aggregate
                    .checked_add(name.len())
                    .ok_or_else(|| make_error("drawing layer text size overflow"))?;
            }
            for layer in &set.layers {
                aggregate = aggregate
                    .checked_add(layer.name.len())
                    .and_then(|size| size.checked_add(layer.title.as_ref().map_or(0, String::len)))
                    .and_then(|size| {
                        size.checked_add(layer.description.as_ref().map_or(0, String::len))
                    })
                    .ok_or_else(|| make_error("drawing layer text size overflow"))?;
            }
            if aggregate > MAX_AGGREGATE_BYTES {
                return invalid("drawing layer text exceeds 16 MiB");
            }
        }
        Ok(())
    }
}

impl crate::OpenDocumentPackage {
    pub fn drawing_layer_sets(&self) -> Result<OdfDrawingLayerSets> {
        let mut sets = parse_part(&self.content_xml()?, OdfDrawingLayerPart::Content)?.sets;
        if let Some(styles) = self.styles_xml()? {
            sets.extend(parse_part(&styles, OdfDrawingLayerPart::Styles)?.sets);
        }
        let result = OdfDrawingLayerSets { sets };
        result.validate()?;
        Ok(result)
    }
}

impl crate::FlatOpenDocument {
    pub fn drawing_layer_sets(&self) -> Result<OdfDrawingLayerSets> {
        parse_drawing_layer_sets(self.xml())
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind {
    None,
    Office,
    Draw,
    Svg,
    Other,
}

#[derive(Clone, Debug)]
enum Container {
    Master(usize),
    Page(usize, Option<String>),
}

#[derive(Clone, Debug)]
struct Frame {
    namespace: NamespaceKind,
    local: String,
    container: Option<Container>,
}

struct ActiveSet {
    parent_depth: usize,
    container_id: usize,
    value: OdfDrawingLayerSet,
}
struct ActiveLayer {
    parent_depth: usize,
    value: OdfDrawingLayer,
    child_order: u8,
}
#[derive(Clone, Copy)]
enum TextKind {
    Title,
    Description,
}
struct ActiveText {
    parent_depth: usize,
    kind: TextKind,
    value: String,
}

type Attributes = HashMap<(NamespaceKind, String), String>;

/// Parses layer sets from a flat ODF document.
pub fn parse_drawing_layer_sets(xml: &str) -> Result<OdfDrawingLayerSets> {
    parse_part(xml, OdfDrawingLayerPart::FlatDocument)
}

fn parse_part(xml: &str, part: OdfDrawingLayerPart) -> Result<OdfDrawingLayerSets> {
    if xml.len() > MAX_XML_BYTES {
        return invalid("drawing layer XML exceeds 64 MiB");
    }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut active_set: Option<ActiveSet> = None;
    let mut active_layer: Option<ActiveLayer> = None;
    let mut active_text: Option<ActiveText> = None;
    let mut seen_containers = HashSet::new();
    let mut next_container_id = 0usize;
    let mut result = OdfDrawingLayerSets::default();

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| make_error(format!("invalid drawing layer XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let attributes = attributes(&mut reader, element)?;
                let container = handle_start(
                    namespace,
                    &local,
                    attributes,
                    part,
                    stack.len(),
                    stack.last(),
                    &mut active_set,
                    &mut active_layer,
                    &mut active_text,
                    &mut seen_containers,
                    &mut next_container_id,
                )?;
                stack.push(Frame {
                    namespace,
                    local,
                    container,
                });
                if stack.len() > MAX_DEPTH {
                    return invalid(format!("drawing layer XML exceeds depth {MAX_DEPTH}"));
                }
            },
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let attributes = attributes(&mut reader, element)?;
                handle_empty(
                    namespace,
                    &local,
                    attributes,
                    part,
                    stack.len(),
                    stack.last(),
                    &mut active_set,
                    &mut active_layer,
                    &mut active_text,
                    &mut seen_containers,
                    &mut next_container_id,
                    &mut result,
                )?;
            },
            Event::End(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                let frame = stack
                    .pop()
                    .ok_or_else(|| make_error("unexpected drawing layer end element"))?;
                if frame.namespace != namespace || frame.local != local {
                    return invalid("drawing layer end element mismatch");
                }
                if active_text
                    .as_ref()
                    .is_some_and(|text| text.parent_depth == stack.len())
                {
                    let text = active_text.take().expect("checked");
                    let layer = active_layer
                        .as_mut()
                        .ok_or_else(|| make_error("layer text has no parent"))?;
                    match text.kind {
                        TextKind::Title => layer.value.title = Some(text.value),
                        TextKind::Description => layer.value.description = Some(text.value),
                    }
                }
                if active_layer
                    .as_ref()
                    .is_some_and(|layer| layer.parent_depth == stack.len())
                {
                    let layer = active_layer.take().expect("checked").value;
                    active_set
                        .as_mut()
                        .ok_or_else(|| make_error("drawing layer has no set"))?
                        .value
                        .layers
                        .push(layer);
                }
                if active_set
                    .as_ref()
                    .is_some_and(|set| set.parent_depth == stack.len())
                {
                    let set = active_set.take().expect("checked");
                    seen_containers.insert(set.container_id);
                    result.sets.push(set.value);
                    if result.sets.len() > MAX_SETS {
                        return invalid(format!("document exceeds {MAX_SETS} drawing layer sets"));
                    }
                }
            },
            Event::Text(ref text) => {
                if let Some(active) = active_text.as_mut() {
                    append_text(
                        &mut active.value,
                        &text
                            .decode()
                            .map_err(|error| make_error(format!("invalid layer text: {error}")))?,
                    )?;
                } else if active_layer.is_some() || active_set.is_some() {
                    let value = text.decode().map_err(|error| {
                        make_error(format!("invalid layer whitespace: {error}"))
                    })?;
                    if !value.trim().is_empty() {
                        return invalid("unexpected character data in drawing layer structure");
                    }
                }
            },
            Event::CData(ref text) => {
                if let Some(active) = active_text.as_mut() {
                    append_text(
                        &mut active.value,
                        &text
                            .decode()
                            .map_err(|error| make_error(format!("invalid layer CDATA: {error}")))?,
                    )?;
                } else if active_layer.is_some() || active_set.is_some() {
                    return invalid("unexpected CDATA in drawing layer structure");
                }
            },
            Event::GeneralRef(ref reference) if active_text.is_some() => {
                let value = if let Some(character) =
                    reference.resolve_char_ref().map_err(|error| {
                        make_error(format!("invalid layer character reference: {error}"))
                    })? {
                    character.to_string()
                } else {
                    match reference
                        .decode()
                        .map_err(|error| {
                            make_error(format!("invalid layer entity reference: {error}"))
                        })?
                        .as_ref()
                    {
                        "amp" => "&".to_owned(),
                        "lt" => "<".to_owned(),
                        "gt" => ">".to_owned(),
                        "apos" => "'".to_owned(),
                        "quot" => "\"".to_owned(),
                        name => {
                            return invalid(format!(
                                "unsupported entity reference '&{name};' in drawing layer metadata"
                            ));
                        },
                    }
                };
                append_text(&mut active_text.as_mut().expect("checked").value, &value)?;
            },
            Event::GeneralRef(_) if active_set.is_some() => {
                return invalid("entity references are not allowed in drawing layer structure");
            },
            Event::DocType(_) => return invalid("DTDs are not allowed in drawing layer XML"),
            Event::PI(_) => {
                return invalid("processing instructions are not allowed in drawing layer XML");
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() || active_set.is_some() || active_layer.is_some() || active_text.is_some()
    {
        return invalid("truncated drawing layer XML");
    }
    result.validate()?;
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn handle_start(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    part: OdfDrawingLayerPart,
    depth: usize,
    parent: Option<&Frame>,
    active_set: &mut Option<ActiveSet>,
    active_layer: &mut Option<ActiveLayer>,
    active_text: &mut Option<ActiveText>,
    seen: &mut HashSet<usize>,
    next_id: &mut usize,
) -> Result<Option<Container>> {
    if active_text.is_some() {
        return invalid("nested elements are not allowed in layer title or description");
    }
    if let Some(layer) = active_layer.as_mut() {
        if depth != layer.parent_depth + 1 {
            return invalid("nested drawing layer content is not allowed");
        }
        start_layer_text(namespace, local, attributes, depth, layer, active_text)?;
        return Ok(None);
    }
    if let Some(set) = active_set.as_mut() {
        if namespace != NamespaceKind::Draw || local != "layer" || depth != set.parent_depth + 1 {
            return invalid("draw:layer-set may contain only direct draw:layer children");
        }
        *active_layer = Some(ActiveLayer {
            parent_depth: depth,
            value: parse_layer(attributes)?,
            child_order: 0,
        });
        return Ok(None);
    }
    if namespace == NamespaceKind::Draw && local == "layer-set" {
        if !attributes.is_empty() {
            return invalid("draw:layer-set does not allow attributes");
        }
        let (container_id, scope) = layer_scope(parent)?;
        if seen.contains(&container_id) {
            return invalid("a layer container may contain only one draw:layer-set");
        }
        *active_set = Some(ActiveSet {
            parent_depth: depth,
            container_id,
            value: OdfDrawingLayerSet {
                part,
                scope,
                layers: Vec::new(),
            },
        });
        return Ok(None);
    }
    if namespace == NamespaceKind::Draw && local == "layer" {
        return invalid("draw:layer must be a direct child of draw:layer-set");
    }
    let container = if namespace == NamespaceKind::Office && local == "master-styles" {
        let id = allocate_id(next_id)?;
        Some(Container::Master(id))
    } else if namespace == NamespaceKind::Draw && local == "page" {
        let name = attributes
            .get(&(NamespaceKind::Draw, "name".to_owned()))
            .cloned();
        if let Some(name) = &name {
            validate_text(name, "draw:page draw:name", false)?;
        }
        let id = allocate_id(next_id)?;
        Some(Container::Page(id, name))
    } else {
        None
    };
    Ok(container)
}

#[allow(clippy::too_many_arguments)]
fn handle_empty(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    part: OdfDrawingLayerPart,
    depth: usize,
    parent: Option<&Frame>,
    active_set: &mut Option<ActiveSet>,
    active_layer: &mut Option<ActiveLayer>,
    active_text: &mut Option<ActiveText>,
    seen: &mut HashSet<usize>,
    next_id: &mut usize,
    result: &mut OdfDrawingLayerSets,
) -> Result<()> {
    if active_text.is_some() {
        return invalid("nested elements are not allowed in layer title or description");
    }
    if let Some(layer) = active_layer.as_mut() {
        if depth != layer.parent_depth + 1 {
            return invalid("nested drawing layer content is not allowed");
        }
        let mut text = None;
        start_layer_text(namespace, local, attributes, depth, layer, &mut text)?;
        let text = text.expect("empty layer text initialized");
        match text.kind {
            TextKind::Title => layer.value.title = Some(String::new()),
            TextKind::Description => layer.value.description = Some(String::new()),
        }
        return Ok(());
    }
    if let Some(set) = active_set.as_mut() {
        if namespace != NamespaceKind::Draw || local != "layer" || depth != set.parent_depth + 1 {
            return invalid("draw:layer-set may contain only direct draw:layer children");
        }
        set.value.layers.push(parse_layer(attributes)?);
        return Ok(());
    }
    if namespace == NamespaceKind::Draw && local == "layer-set" {
        if !attributes.is_empty() {
            return invalid("draw:layer-set does not allow attributes");
        }
        let (container_id, scope) = layer_scope(parent)?;
        if !seen.insert(container_id) {
            return invalid("a layer container may contain only one draw:layer-set");
        }
        result.sets.push(OdfDrawingLayerSet {
            part,
            scope,
            layers: Vec::new(),
        });
        return Ok(());
    }
    if namespace == NamespaceKind::Draw && local == "layer" {
        return invalid("draw:layer must be a direct child of draw:layer-set");
    }
    if (namespace == NamespaceKind::Office && local == "master-styles")
        || (namespace == NamespaceKind::Draw && local == "page")
    {
        let _ = allocate_id(next_id)?;
    }
    Ok(())
}

fn start_layer_text(
    namespace: NamespaceKind,
    local: &str,
    attributes: Attributes,
    depth: usize,
    layer: &mut ActiveLayer,
    active: &mut Option<ActiveText>,
) -> Result<()> {
    if !attributes.is_empty() {
        return invalid("layer title and description do not allow attributes");
    }
    let (kind, order) = match (namespace, local) {
        (NamespaceKind::Svg, "title") => (TextKind::Title, 1),
        (NamespaceKind::Svg, "desc") => (TextKind::Description, 2),
        _ => return invalid("draw:layer may contain only svg:title followed by svg:desc"),
    };
    if order <= layer.child_order {
        return invalid("duplicate or out-of-order layer title/description");
    }
    layer.child_order = order;
    *active = Some(ActiveText {
        parent_depth: depth,
        kind,
        value: String::new(),
    });
    Ok(())
}

fn parse_layer(mut attributes: Attributes) -> Result<OdfDrawingLayer> {
    let name = attributes
        .remove(&(NamespaceKind::Draw, "name".to_owned()))
        .ok_or_else(|| make_error("draw:layer requires draw:name"))?;
    let protected = attributes
        .remove(&(NamespaceKind::Draw, "protected".to_owned()))
        .map(|value| parse_bool(&value))
        .transpose()?;
    let display = attributes
        .remove(&(NamespaceKind::Draw, "display".to_owned()))
        .map(|value| OdfDrawingLayerDisplay::parse(&value))
        .transpose()?;
    if let Some(((namespace, local), _)) = attributes.into_iter().next() {
        return invalid(format!(
            "unsupported {:?} layer attribute '{local}'",
            namespace
        ));
    }
    let layer = OdfDrawingLayer {
        name,
        protected,
        display,
        title: None,
        description: None,
    };
    layer.validate()?;
    Ok(layer)
}

fn layer_scope(parent: Option<&Frame>) -> Result<(usize, OdfDrawingLayerScope)> {
    match parent.and_then(|frame| frame.container.as_ref()) {
        Some(Container::Master(id)) => Ok((*id, OdfDrawingLayerScope::MasterStyles)),
        Some(Container::Page(id, name)) => {
            Ok((*id, OdfDrawingLayerScope::Page { name: name.clone() }))
        },
        None => {
            invalid("draw:layer-set must be a direct child of office:master-styles or draw:page")
        },
    }
}

fn attributes<R: BufRead>(
    reader: &mut NsReader<R>,
    element: &BytesStart<'_>,
) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute
            .map_err(|error| make_error(format!("invalid drawing layer attribute: {error}")))?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver_mut().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        reject_spoofed_name(namespace, &local)?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| make_error(format!("invalid drawing layer attribute value: {error}")))?
            .into_owned();
        validate_text(&value, &local, true)?;
        if result.insert((namespace, local.clone()), value).is_some() {
            return invalid(format!("duplicate drawing layer attribute '{local}'"));
        }
    }
    Ok(result)
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) => match namespace.as_ref() {
            OFFICE_NS => Ok(NamespaceKind::Office),
            DRAW_NS => Ok(NamespaceKind::Draw),
            SVG_NS => Ok(NamespaceKind::Svg),
            _ => Ok(NamespaceKind::Other),
        },
        ResolveResult::Unknown(prefix) => invalid(format!(
            "unknown XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        )),
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if matches!(local, "layer-set" | "layer") && namespace != NamespaceKind::Draw {
        return invalid(format!("spoofed draw:{local} element namespace"));
    }
    Ok(())
}

fn allocate_id(next: &mut usize) -> Result<usize> {
    let id = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| make_error("drawing layer container count overflow"))?;
    Ok(id)
}
fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid ODF boolean '{value}'")),
    }
}
fn decode(value: &[u8], name: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_owned)
        .map_err(|error| make_error(format!("invalid UTF-8 {name}: {error}")))
}
fn append_text(target: &mut String, value: &str) -> Result<()> {
    if target.len().saturating_add(value.len()) > MAX_VALUE_BYTES {
        return invalid(format!(
            "drawing layer text exceeds {MAX_VALUE_BYTES} bytes"
        ));
    }
    target.push_str(value);
    Ok(())
}
fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> {
    if !allow_empty && value.is_empty() {
        return invalid(format!("{name} cannot be empty"));
    }
    if value.len() > MAX_VALUE_BYTES {
        return invalid(format!("{name} exceeds {MAX_VALUE_BYTES} bytes"));
    }
    if value.chars().any(
        |character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}'),
    ) {
        return invalid(format!("{name} contains invalid XML characters"));
    }
    Ok(())
}
fn escape_attribute(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\r' => output.push_str("&#13;"),
            '\n' => output.push_str("&#10;"),
            '\t' => output.push_str("&#9;"),
            _ => output.push(character),
        }
    }
}
fn escape_text(output: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\r' => output.push_str("&#13;"),
            _ => output.push(character),
        }
    }
}
fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

#[cfg(test)]
mod tests {
    use super::*;

    const PREFIX: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">"#;

    #[test]
    fn parses_scoped_layers_and_round_trips_fragments() {
        let xml = format!(
            r#"{PREFIX}<office:master-styles><draw:layer-set><draw:layer draw:name="layout" draw:protected="1" draw:display="screen"><svg:title>A &amp; B</svg:title><svg:desc>Visible</svg:desc></draw:layer><draw:layer draw:name="print" draw:display="printer"/><draw:layer draw:name="hidden" draw:display="none"/></draw:layer-set></office:master-styles><draw:page draw:name="Page 1"><draw:layer-set><draw:layer draw:name="default" draw:display="always"/></draw:layer-set></draw:page></office:document>"#
        );
        let parsed = parse_drawing_layer_sets(&xml).unwrap();
        assert_eq!(parsed.sets.len(), 2);
        assert_eq!(parsed.sets[0].layers[0].title.as_deref(), Some("A & B"));
        assert!(parsed.sets[0].layers[0].effective_protected());
        assert_eq!(
            parsed.sets[1].scope,
            OdfDrawingLayerScope::Page {
                name: Some("Page 1".into())
            }
        );
        let fragment = parsed.sets[0].to_xml_fragment().unwrap();
        let reparsed = parse_drawing_layer_sets(&format!(
            r#"{PREFIX}<office:master-styles>{fragment}</office:master-styles></office:document>"#
        ))
        .unwrap();
        assert_eq!(reparsed.sets[0].layers, parsed.sets[0].layers);
    }

    #[test]
    fn rejects_invalid_layer_grammar() {
        for body in [
            r#"<draw:layer-set><draw:layer draw:name="x"/><draw:layer draw:name="x"/></draw:layer-set>"#,
            r#"<draw:layer-set><draw:layer draw:name="x"><svg:desc/><svg:title/></draw:layer></draw:layer-set>"#,
            r#"<draw:layer-set><draw:layer draw:name="x" draw:protected="yes"/></draw:layer-set>"#,
            r#"<draw:layer-set bad="x"/>"#,
            r#"<draw:layer-set><draw:layer draw:name="x"><svg:title><svg:desc/></svg:title></draw:layer></draw:layer-set>"#,
        ] {
            let xml = format!(
                r#"{PREFIX}<office:master-styles>{body}</office:master-styles></office:document>"#
            );
            assert!(parse_drawing_layer_sets(&xml).is_err(), "accepted {body}");
        }
        assert!(
            parse_drawing_layer_sets(&format!(r#"{PREFIX}<draw:layer-set/></office:document>"#))
                .is_err()
        );
        assert!(parse_drawing_layer_sets(&format!(r#"{PREFIX}<office:master-styles><draw:layer-set/><draw:layer-set/></office:master-styles></office:document>"#)).is_err());
        assert!(parse_drawing_layer_sets("<!DOCTYPE x><x/>").is_err());
    }

    #[test]
    fn parses_libreoffice_default_layer_set_when_fixture_is_available() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/libreoffice-core/xmloff/qa/unit/data/theme.fodp");
        let Ok(xml) = std::fs::read_to_string(path) else {
            return;
        };
        let parsed = parse_drawing_layer_sets(&xml).unwrap();
        let set = parsed
            .sets
            .iter()
            .find(|set| matches!(set.scope, OdfDrawingLayerScope::MasterStyles))
            .unwrap();
        assert_eq!(
            set.layers
                .iter()
                .map(|layer| layer.name.as_str())
                .collect::<Vec<_>>(),
            [
                "layout",
                "background",
                "backgroundobjects",
                "controls",
                "measurelines"
            ]
        );
    }
}
