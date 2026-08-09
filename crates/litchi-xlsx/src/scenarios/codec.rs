//! Bounded `SpreadsheetML` scenario XML codec.

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::writer::Writer;

use crate::error::{Result, invalid};
use litchi_ooxml_common::mce::process_str;

use super::model::{
    CellReference, ChildOrder, Collection, Conformance, InputCell, MAX_DEPTH, MAX_EVENTS,
    MAX_INPUT_CELLS, MAX_SCENARIOS, MAX_SQREF_ITEMS, MAX_UNKNOWN_BYTES, MAX_XML_BYTES,
    NamespaceBinding, OpaqueFields, RangeReference, STRICT_MAIN, Scenario, TRANSITIONAL_MAIN,
    UnknownAttribute, UnknownElement, checked_xstring,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    Collection,
    Scenario,
    InputCells,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Unbound,
    Main,
    Other,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum CaptureOwner {
    Collection,
    Scenario,
    InputCell,
}

struct Capture {
    owner: CaptureOwner,
    depth: usize,
    writer: Writer<Vec<u8>>,
}

/// Parses the direct worksheet `scenarios` child after applying shared MCE processing.
pub fn parse_worksheet_scenarios(xml: &[u8]) -> Result<Option<Collection>> {
    if xml.len() > MAX_XML_BYTES {
        return Err(invalid("worksheet XML exceeds safety limit"));
    }
    let source = std::str::from_utf8(xml)
        .map_err(|error| invalid(format!("worksheet XML is not UTF-8: {error}")))?;
    let processed = process_str(source)?;
    if processed.len() > MAX_XML_BYTES {
        return Err(invalid("processed worksheet XML exceeds safety limit"));
    }
    let mut reader = NsReader::from_reader(processed.as_bytes());
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut scopes = Vec::new();
    let mut state: Option<CollectionBuilder> = None;
    let mut capture: Option<Capture> = None;
    let mut seen_scenarios = false;
    let mut depth = 0usize;
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
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid worksheet XML: {error}")))?;
        let namespace = namespace_kind(resolved)?;

        if let Some(active) = capture.as_mut() {
            let complete = match &event {
                Event::Start(_) => {
                    active.depth = active
                        .depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("unknown scenario element depth overflow"))?;
                    if active.depth > MAX_DEPTH {
                        return Err(invalid("unknown scenario element nesting is too deep"));
                    }
                    depth = depth
                        .checked_add(1)
                        .ok_or_else(|| invalid("worksheet XML depth overflow"))?;
                    false
                },
                Event::End(_) => {
                    active.depth = active
                        .depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("unknown scenario element depth underflow"))?;
                    depth = depth
                        .checked_sub(1)
                        .ok_or_else(|| invalid("worksheet XML depth underflow"))?;
                    active.depth == 0
                },
                Event::Eof => return Err(invalid("unterminated unknown scenario element")),
                _ => false,
            };
            active
                .writer
                .write_event(event.clone())
                .map_err(|error| invalid(format!("invalid unknown scenario element: {error}")))?;
            if active.writer.get_ref().len() > MAX_UNKNOWN_BYTES {
                return Err(invalid("unknown scenario element exceeds safety limit"));
            }
            if complete {
                let active = capture
                    .take()
                    .ok_or_else(|| invalid("missing unknown scenario capture"))?;
                let element = UnknownElement::from_parts(active.writer.into_inner(), Vec::new())?;
                attach_unknown(&mut state, active.owner, element)?;
            }
            buffer.clear();
            continue;
        }
        match event {
            Event::Start(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if scopes.is_empty() && root_seen {
                    return Err(invalid("worksheet XML contains multiple roots"));
                }
                let next_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet XML depth overflow"))?;
                if next_depth > MAX_DEPTH {
                    return Err(invalid("worksheet XML nesting is too deep"));
                }
                if let Some(owner) = capture_owner(
                    scopes.last().copied(),
                    namespace,
                    element.local_name().as_ref(),
                ) {
                    capture = Some(start_capture(&reader, &element, owner)?);
                    depth = next_depth;
                    buffer.clear();
                    continue;
                }
                let scope = begin_element(
                    &reader,
                    &element,
                    namespace,
                    scopes.last().copied(),
                    &mut state,
                    &mut seen_scenarios,
                )?;
                if scopes.is_empty() {
                    root_seen = true;
                }
                depth = next_depth;
                scopes.push(scope);
            },
            Event::Empty(element) => {
                if root_closed {
                    return Err(invalid("worksheet XML contains content after root"));
                }
                if scopes.is_empty() {
                    if root_seen {
                        return Err(invalid("worksheet XML contains multiple roots"));
                    }
                    let scope = begin_element(
                        &reader,
                        &element,
                        namespace,
                        None,
                        &mut state,
                        &mut seen_scenarios,
                    )?;
                    end_scope(scope, &mut state)?;
                    root_seen = true;
                    root_closed = true;
                } else {
                    if let Some(owner) = capture_owner(
                        scopes.last().copied(),
                        namespace,
                        element.local_name().as_ref(),
                    ) {
                        let unknown = capture_empty(&reader, &element)?;
                        attach_unknown(&mut state, owner, unknown)?;
                        buffer.clear();
                        continue;
                    }
                    let scope = begin_element(
                        &reader,
                        &element,
                        namespace,
                        scopes.last().copied(),
                        &mut state,
                        &mut seen_scenarios,
                    )?;
                    end_scope(scope, &mut state)?;
                }
            },
            Event::End(element) => {
                let scope = scopes
                    .pop()
                    .ok_or_else(|| invalid("unexpected worksheet end element"))?;
                match scope {
                    Scope::Worksheet => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"worksheet"
                        {
                            return Err(invalid("mismatched worksheet end element"));
                        }
                        root_closed = true;
                    },
                    Scope::Collection => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"scenarios"
                        {
                            return Err(invalid("mismatched scenarios end element"));
                        }
                    },
                    Scope::Scenario => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"scenario"
                        {
                            return Err(invalid("mismatched scenario end element"));
                        }
                    },
                    Scope::InputCells => {
                        if namespace != NamespaceKind::Main
                            || element.local_name().as_ref() != b"inputCells"
                        {
                            return Err(invalid("mismatched inputCells end element"));
                        }
                    },
                    Scope::Other => {},
                }
                end_scope(scope, &mut state)?;
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("worksheet XML depth underflow"))?;
            },
            Event::Text(text)
                if matches!(
                    scopes.last(),
                    Some(Scope::Collection | Scope::Scenario | Scope::InputCells)
                ) && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("scenarios family cannot contain text"));
            },
            Event::Text(text)
                if matches!(scopes.last(), Some(Scope::Worksheet))
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("worksheet cannot contain direct text"));
            },
            Event::CData(text)
                if matches!(
                    scopes.last(),
                    Some(Scope::Collection | Scope::Scenario | Scope::InputCells)
                ) && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("scenarios family cannot contain CDATA"));
            },
            Event::CData(_) if matches!(scopes.last(), Some(Scope::Worksheet)) => {
                return Err(invalid("worksheet cannot contain direct CDATA"));
            },
            Event::Text(text)
                if scopes.is_empty() && !text.as_ref().iter().all(u8::is_ascii_whitespace) =>
            {
                return Err(invalid("worksheet XML text is outside root"));
            },
            Event::CData(_) if scopes.is_empty() => {
                return Err(invalid("worksheet XML CDATA is outside root"));
            },
            Event::Decl(_) => {
                if root_seen || declaration_seen {
                    return Err(invalid("invalid worksheet XML declaration position"));
                }
                declaration_seen = true;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen || !root_closed || depth != 0 || !scopes.is_empty() {
        return Err(invalid("unterminated worksheet XML"));
    }
    state.map(finish_builder).transpose()
}

#[derive(Default)]
struct CollectionBuilder {
    current: Option<u32>,
    show: Option<u32>,
    ranges: Vec<RangeReference>,
    scenarios: Vec<Scenario>,
    open_scenario: Option<ScenarioBuilder>,
    opaque: OpaqueFields,
}

#[derive(Default)]
struct ScenarioBuilder {
    name: Option<String>,
    locked: Option<bool>,
    hidden: Option<bool>,
    count: Option<u32>,
    user: Option<String>,
    comment: Option<String>,
    input_cells: Vec<InputCell>,
    opaque: OpaqueFields,
}

fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    state: &mut Option<CollectionBuilder>,
    seen_scenarios: &mut bool,
) -> Result<Scope> {
    let local = element.local_name();
    let local = local.as_ref();
    let main = namespace == NamespaceKind::Main;
    match parent {
        None => {
            if !main || local != b"worksheet" {
                return Err(invalid("expected SpreadsheetML worksheet root"));
            }
            Ok(Scope::Worksheet)
        },
        Some(Scope::Worksheet) => {
            if local != b"scenarios" {
                return Ok(Scope::Other);
            }
            if !main {
                return Err(invalid("spoofed scenarios element namespace"));
            }
            if *seen_scenarios {
                return Err(invalid("duplicate worksheet scenarios element"));
            }
            *seen_scenarios = true;
            *state = Some(parse_scenarios_attributes(reader, element)?);
            Ok(Scope::Collection)
        },
        Some(Scope::Collection) => {
            if local != b"scenario" || !main {
                return Err(invalid(if local == b"scenario" {
                    "spoofed scenario element namespace"
                } else {
                    "unknown scenarios child element"
                }));
            }
            let builder = state
                .as_mut()
                .ok_or_else(|| invalid("missing scenarios state"))?;
            if builder.open_scenario.is_some() {
                return Err(invalid("nested scenario element"));
            }
            if builder.scenarios.len() >= MAX_SCENARIOS {
                return Err(invalid(format!(
                    "scenarios exceeds safety limit {MAX_SCENARIOS}"
                )));
            }
            builder.open_scenario = Some(parse_scenario_attributes(reader, element)?);
            Ok(Scope::Scenario)
        },
        Some(Scope::Scenario) => {
            if local == b"inputCells" && main {
                let builder = state
                    .as_mut()
                    .and_then(|value| value.open_scenario.as_mut())
                    .ok_or_else(|| invalid("missing scenario state"))?;
                if builder.input_cells.len() >= MAX_INPUT_CELLS {
                    return Err(invalid(format!(
                        "scenario inputCells exceeds safety limit {MAX_INPUT_CELLS}"
                    )));
                }
                let index = builder.input_cells.len();
                builder
                    .input_cells
                    .push(parse_input_cell_attributes(reader, element)?);
                builder.opaque.push_known(index)?;
                Ok(Scope::InputCells)
            } else if local == b"extLst" && main {
                // Extension markup is inert and not interpreted; skip the subtree.
                Ok(Scope::Other)
            } else {
                Err(invalid(if local == b"inputCells" || local == b"extLst" {
                    "spoofed scenario child element namespace"
                } else {
                    "unknown scenario child element"
                }))
            }
        },
        Some(Scope::InputCells) => Err(invalid("inputCells must be a leaf element")),
        Some(Scope::Other) => Ok(Scope::Other),
    }
}

fn end_scope(scope: Scope, state: &mut Option<CollectionBuilder>) -> Result<()> {
    if scope != Scope::Scenario {
        return Ok(());
    }
    let builder = state
        .as_mut()
        .ok_or_else(|| invalid("missing scenarios state"))?;
    let scenario = builder
        .open_scenario
        .take()
        .ok_or_else(|| invalid("missing scenario state"))?;
    let name = scenario
        .name
        .ok_or_else(|| invalid("scenario requires name"))?;
    let index = builder.scenarios.len();
    builder.scenarios.push(Scenario {
        name,
        locked: scenario.locked.unwrap_or(false),
        hidden: scenario.hidden.unwrap_or(false),
        count: scenario.count,
        user: scenario.user,
        comment: scenario.comment,
        input_cells: scenario.input_cells,
        opaque: optional_opaque(scenario.opaque),
    });
    builder.opaque.push_known(index)?;
    Ok(())
}

fn finish_builder(builder: CollectionBuilder) -> Result<Collection> {
    if builder.scenarios.is_empty() {
        return Err(invalid("scenarios requires at least one scenario"));
    }
    Ok(Collection {
        current: builder.current,
        show: builder.show,
        ranges: builder.ranges,
        scenarios: builder.scenarios,
        opaque: optional_opaque(builder.opaque),
    })
}

fn parse_scenarios_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<CollectionBuilder> {
    let mut value = CollectionBuilder::default();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid scenarios attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid scenarios attribute value: {error}")))?;
        if !matches!(namespace, ResolveResult::Unbound) {
            preserve_attribute(&mut value.opaque, reader, attribute, namespace, text)?;
            continue;
        }
        match local.as_ref() {
            b"current" => set_once(
                &mut value.current,
                parse_u32(&text, "scenarios current")?,
                "current",
            )?,
            b"show" => set_once(&mut value.show, parse_u32(&text, "scenarios show")?, "show")?,
            b"sqref" => {
                if !value.ranges.is_empty() {
                    return Err(invalid("duplicate scenarios sqref attribute"));
                }
                value.ranges = parse_sqref(&text)?;
            },
            _ => preserve_attribute(&mut value.opaque, reader, attribute, namespace, text)?,
        }
    }
    Ok(value)
}

fn parse_scenario_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ScenarioBuilder> {
    let mut value = ScenarioBuilder::default();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid scenario attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid scenario attribute value: {error}")))?
            .into_owned();
        if !matches!(namespace, ResolveResult::Unbound) {
            preserve_attribute(&mut value.opaque, reader, attribute, namespace, text)?;
            continue;
        }
        match local.as_ref() {
            b"name" => set_once(
                &mut value.name,
                checked_xstring(text, "scenario name")?,
                "name",
            )?,
            b"locked" => set_once(&mut value.locked, parse_bool(&text, "locked")?, "locked")?,
            b"hidden" => set_once(&mut value.hidden, parse_bool(&text, "hidden")?, "hidden")?,
            b"count" => set_once(
                &mut value.count,
                parse_u32(&text, "scenario count")?,
                "count",
            )?,
            b"user" => set_once(
                &mut value.user,
                checked_xstring(text, "scenario user")?,
                "user",
            )?,
            b"comment" => set_once(
                &mut value.comment,
                checked_xstring(text, "scenario comment")?,
                "comment",
            )?,
            _ => preserve_attribute(&mut value.opaque, reader, attribute, namespace, text)?,
        }
    }
    Ok(value)
}

fn parse_input_cell_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<InputCell> {
    let mut reference = None;
    let mut deleted = None;
    let mut undone = None;
    let mut input_value = None;
    let mut number_format_id = None;
    let mut opaque = OpaqueFields::default();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid inputCells attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid inputCells attribute value: {error}")))?
            .into_owned();
        if !matches!(namespace, ResolveResult::Unbound) {
            preserve_attribute(&mut opaque, reader, attribute, namespace, text)?;
            continue;
        }
        match local.as_ref() {
            b"r" => set_once(&mut reference, CellReference::new(text)?, "r")?,
            b"deleted" => set_once(&mut deleted, parse_bool(&text, "deleted")?, "deleted")?,
            b"undone" => set_once(&mut undone, parse_bool(&text, "undone")?, "undone")?,
            b"val" => set_once(
                &mut input_value,
                checked_xstring(text, "inputCells val")?,
                "val",
            )?,
            b"numFmtId" => set_once(
                &mut number_format_id,
                parse_u32(&text, "inputCells numFmtId")?,
                "numFmtId",
            )?,
            _ => preserve_attribute(&mut opaque, reader, attribute, namespace, text)?,
        }
    }
    Ok(InputCell {
        reference: reference.ok_or_else(|| invalid("inputCells requires r"))?,
        deleted: deleted.unwrap_or(false),
        undone: undone.unwrap_or(false),
        value: input_value.ok_or_else(|| invalid("inputCells requires val"))?,
        number_format_id,
        opaque: optional_opaque(opaque),
    })
}

fn preserve_attribute(
    opaque: &mut OpaqueFields,
    reader: &NsReader<&[u8]>,
    attribute: quick_xml::events::attributes::Attribute<'_>,
    namespace: ResolveResult<'_>,
    value: impl Into<String>,
) -> Result<()> {
    let qualified_name = String::from_utf8(attribute.key.as_ref().to_vec()).map_err(|error| {
        invalid(format!(
            "unknown scenario attribute name is not UTF-8: {error}"
        ))
    })?;
    let namespace = match namespace {
        ResolveResult::Unbound => None,
        ResolveResult::Bound(uri) => {
            let prefix = attribute
                .key
                .as_ref()
                .split(|byte| *byte == b':')
                .next()
                .filter(|prefix| *prefix != attribute.key.as_ref())
                .ok_or_else(|| invalid("namespaced scenario attribute has no prefix"))?;
            Some(NamespaceBinding::new(
                String::from_utf8(prefix.to_vec()).map_err(|error| {
                    invalid(format!("scenario attribute prefix is not UTF-8: {error}"))
                })?,
                String::from_utf8(uri.as_ref().to_vec()).map_err(|error| {
                    invalid(format!(
                        "scenario attribute namespace is not UTF-8: {error}"
                    ))
                })?,
            ))
        },
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound XML namespace prefix {}",
                String::from_utf8_lossy(&prefix)
            )));
        },
    };
    let _ = reader;
    opaque.push_attribute(UnknownAttribute::from_decoded(
        qualified_name,
        value.into(),
        namespace,
    )?)
}

fn capture_owner(
    parent: Option<Scope>,
    namespace: NamespaceKind,
    local: &[u8],
) -> Option<CaptureOwner> {
    match parent {
        Some(Scope::Collection) if !(namespace == NamespaceKind::Main && local == b"scenario") => {
            Some(CaptureOwner::Collection)
        },
        Some(Scope::Scenario) if !(namespace == NamespaceKind::Main && local == b"inputCells") => {
            Some(CaptureOwner::Scenario)
        },
        Some(Scope::InputCells) => Some(CaptureOwner::InputCell),
        _ => None,
    }
}

fn start_capture(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    owner: CaptureOwner,
) -> Result<Capture> {
    let mut root = element.to_owned();
    add_namespace_declarations(reader, &mut root)?;
    let mut writer = Writer::new(Vec::new());
    writer
        .write_event(Event::Start(root))
        .map_err(|error| invalid(format!("invalid unknown scenario element: {error}")))?;
    Ok(Capture {
        owner,
        depth: 1,
        writer,
    })
}

fn capture_empty(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<UnknownElement> {
    let mut root = element.to_owned();
    add_namespace_declarations(reader, &mut root)?;
    let mut writer = Writer::new(Vec::new());
    writer
        .write_event(Event::Empty(root))
        .map_err(|error| invalid(format!("invalid unknown scenario element: {error}")))?;
    UnknownElement::from_parts(writer.into_inner(), Vec::new())
}

fn add_namespace_declarations(
    reader: &NsReader<&[u8]>,
    element: &mut BytesStart<'static>,
) -> Result<()> {
    let mut declared = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute
            .map_err(|error| invalid(format!("invalid unknown scenario attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            declared.push(attribute.key.as_ref().to_vec());
        }
    }
    for (prefix, namespace) in reader.resolver().bindings() {
        let key = match prefix {
            PrefixDeclaration::Default => b"xmlns".to_vec(),
            PrefixDeclaration::Named(prefix) => {
                let mut key = b"xmlns:".to_vec();
                key.extend_from_slice(prefix);
                key
            },
        };
        if declared.iter().any(|value| value == &key) {
            continue;
        }
        element.push_attribute((key.as_slice(), namespace.into_inner()));
        declared.push(key);
    }
    Ok(())
}

fn attach_unknown(
    state: &mut Option<CollectionBuilder>,
    owner: CaptureOwner,
    element: UnknownElement,
) -> Result<()> {
    let builder = state
        .as_mut()
        .ok_or_else(|| invalid("unknown scenario element has no collection"))?;
    match owner {
        CaptureOwner::Collection => builder.opaque.push_element(element).map(|_| ()),
        CaptureOwner::Scenario => builder
            .open_scenario
            .as_mut()
            .ok_or_else(|| invalid("unknown scenario element has no scenario"))?
            .opaque
            .push_element(element)
            .map(|_| ()),
        CaptureOwner::InputCell => builder
            .open_scenario
            .as_mut()
            .ok_or_else(|| invalid("unknown scenario element has no scenario"))?
            .input_cells
            .last_mut()
            .ok_or_else(|| invalid("unknown scenario element has no input cell"))?
            .opaque_mut()
            .push_element(element)
            .map(|_| ()),
    }
}

fn optional_opaque(opaque: OpaqueFields) -> Option<Box<OpaqueFields>> {
    opaque.has_unknown().then(|| Box::new(opaque))
}

/// Serializes one canonical, namespace-complete `scenarios` fragment.
pub fn write_worksheet_scenarios(value: &Collection, conformance: Conformance) -> Result<String> {
    if value.scenarios().is_empty() {
        return Err(invalid("scenarios requires at least one scenario"));
    }
    if value.scenarios().len() > MAX_SCENARIOS {
        return Err(invalid(format!(
            "scenarios exceeds safety limit {MAX_SCENARIOS}"
        )));
    }
    let mut xml = BoundedXml::new();
    xml.push_str("<scenarios xmlns=\"")?;
    xml.push_str(conformance.main_namespace())?;
    xml.push_char('"')?;
    if let Some(current) = value.current() {
        write_u32_attribute(&mut xml, "current", current)?;
    }
    if let Some(show) = value.show() {
        write_u32_attribute(&mut xml, "show", show)?;
    }
    if !value.ranges().is_empty() {
        xml.push_str(" sqref=\"")?;
        for (index, range) in value.ranges().iter().enumerate() {
            if index > 0 {
                xml.push_char(' ')?;
            }
            xml.push_str(range.as_str())?;
        }
        xml.push_char('"')?;
    }
    write_unknown_attributes(&mut xml, value.opaque.as_deref())?;
    xml.push_char('>')?;
    if let Some(opaque) = value.opaque.as_deref() {
        if opaque.order.is_empty() {
            for scenario in value.scenarios() {
                write_scenario(&mut xml, scenario)?;
            }
        } else {
            for child in &opaque.order {
                match *child {
                    ChildOrder::Known(index) => write_scenario(
                        &mut xml,
                        value
                            .scenarios()
                            .get(index)
                            .ok_or_else(|| invalid("invalid scenario child order"))?,
                    )?,
                    ChildOrder::Unknown(index) => write_unknown_element(
                        &mut xml,
                        opaque
                            .elements
                            .get(index)
                            .ok_or_else(|| invalid("invalid unknown scenario child order"))?,
                    )?,
                }
            }
        }
    } else {
        for scenario in value.scenarios() {
            write_scenario(&mut xml, scenario)?;
        }
    }
    xml.push_str("</scenarios>")?;
    xml.finish()
}

fn write_scenario(xml: &mut BoundedXml, scenario: &Scenario) -> Result<()> {
    xml.push_str("<scenario")?;
    write_attribute(xml, "name", scenario.name())?;
    write_true_attribute(xml, "locked", scenario.locked())?;
    write_true_attribute(xml, "hidden", scenario.hidden())?;
    if let Some(count) = scenario.count() {
        write_u32_attribute(xml, "count", count)?;
    }
    if let Some(user) = scenario.user() {
        write_attribute(xml, "user", user)?;
    }
    if let Some(comment) = scenario.comment() {
        write_attribute(xml, "comment", comment)?;
    }
    write_unknown_attributes(xml, scenario.opaque.as_deref())?;
    let has_children = !scenario.input_cells().is_empty()
        || scenario
            .opaque
            .as_deref()
            .is_some_and(|opaque| !opaque.elements.is_empty());
    if !has_children {
        xml.push_str("/>")?;
        return Ok(());
    }
    xml.push_char('>')?;
    if let Some(opaque) = scenario.opaque.as_deref() {
        if opaque.order.is_empty() {
            for cell in scenario.input_cells() {
                write_input_cell(xml, cell)?;
            }
        } else {
            for child in &opaque.order {
                match *child {
                    ChildOrder::Known(index) => write_input_cell(
                        xml,
                        scenario
                            .input_cells()
                            .get(index)
                            .ok_or_else(|| invalid("invalid input-cell child order"))?,
                    )?,
                    ChildOrder::Unknown(index) => write_unknown_element(
                        xml,
                        opaque
                            .elements
                            .get(index)
                            .ok_or_else(|| invalid("invalid unknown scenario child order"))?,
                    )?,
                }
            }
        }
    } else {
        for cell in scenario.input_cells() {
            write_input_cell(xml, cell)?;
        }
    }
    xml.push_str("</scenario>")?;
    Ok(())
}

fn write_input_cell(xml: &mut BoundedXml, cell: &InputCell) -> Result<()> {
    xml.push_str("<inputCells")?;
    write_attribute(xml, "r", cell.reference().as_str())?;
    write_true_attribute(xml, "deleted", cell.deleted())?;
    write_true_attribute(xml, "undone", cell.undone())?;
    write_attribute(xml, "val", cell.value())?;
    if let Some(number_format_id) = cell.number_format_id() {
        write_u32_attribute(xml, "numFmtId", number_format_id)?;
    }
    write_unknown_attributes(xml, cell.opaque.as_deref())?;
    let has_children = cell
        .opaque
        .as_deref()
        .is_some_and(|opaque| !opaque.elements.is_empty());
    if !has_children {
        xml.push_str("/>")?;
        return Ok(());
    }
    xml.push_char('>')?;
    if let Some(opaque) = cell.opaque.as_deref() {
        for child in &opaque.order {
            match *child {
                ChildOrder::Known(_) => {
                    return Err(invalid("inputCells contains an invalid known child"));
                },
                ChildOrder::Unknown(index) => write_unknown_element(
                    xml,
                    opaque
                        .elements
                        .get(index)
                        .ok_or_else(|| invalid("invalid unknown input-cell child order"))?,
                )?,
            }
        }
    }
    xml.push_str("</inputCells>")?;
    Ok(())
}

fn write_unknown_attributes(xml: &mut BoundedXml, opaque: Option<&OpaqueFields>) -> Result<()> {
    let Some(opaque) = opaque else {
        return Ok(());
    };
    let mut declared = Vec::<String>::new();
    for attribute in &opaque.attributes {
        if let Some(namespace) = attribute.namespace.as_ref()
            && namespace.prefix.as_ref() != "xml"
            && !declared
                .iter()
                .any(|prefix| prefix == namespace.prefix.as_ref())
        {
            let declaration = format!("xmlns:{}", namespace.prefix);
            write_attribute(xml, &declaration, &namespace.uri)?;
            declared.push(namespace.prefix.to_string());
        }
        write_attribute(xml, attribute.name(), attribute.value())?;
    }
    Ok(())
}

fn write_unknown_element(xml: &mut BoundedXml, element: &UnknownElement) -> Result<()> {
    xml.push_bytes(element.as_xml())
}

fn write_true_attribute(xml: &mut BoundedXml, name: &str, value: bool) -> Result<()> {
    if value {
        xml.push_char(' ')?;
        xml.push_str(name)?;
        xml.push_str("=\"1\"")?;
    }
    Ok(())
}

fn write_u32_attribute(xml: &mut BoundedXml, name: &str, value: u32) -> Result<()> {
    xml.push_char(' ')?;
    xml.push_str(name)?;
    xml.push_str("=\"")?;
    let value = value.to_string();
    xml.push_str(&value)?;
    xml.push_char('"')?;
    Ok(())
}

fn write_attribute(xml: &mut BoundedXml, name: &str, value: &str) -> Result<()> {
    xml.push_char(' ')?;
    xml.push_str(name)?;
    xml.push_str("=\"")?;
    let mut plain_start = 0;
    for (index, character) in value.char_indices() {
        let escaped = match character {
            '&' => Some("&amp;"),
            '<' => Some("&lt;"),
            '>' => Some("&gt;"),
            '"' => Some("&quot;"),
            '\'' => Some("&apos;"),
            _ => None,
        };
        if let Some(escaped) = escaped {
            xml.push_str(&value[plain_start..index])?;
            xml.push_str(escaped)?;
            plain_start = index + character.len_utf8();
        }
    }
    xml.push_str(&value[plain_start..])?;
    xml.push_char('"')?;
    Ok(())
}

/// Fallible XML sink that checks the serialized part limit before each append.
struct BoundedXml {
    bytes: Vec<u8>,
}

impl BoundedXml {
    fn new() -> Self {
        Self { bytes: Vec::new() }
    }

    fn push_str(&mut self, value: &str) -> Result<()> {
        self.push_bytes(value.as_bytes())
    }

    fn push_char(&mut self, value: char) -> Result<()> {
        let mut encoded = [0; 4];
        let length = value.encode_utf8(&mut encoded).len();
        self.push_bytes(&encoded[..length])
    }

    fn push_bytes(&mut self, value: &[u8]) -> Result<()> {
        let length = self
            .bytes
            .len()
            .checked_add(value.len())
            .ok_or_else(|| invalid("serialized scenarios XML length overflows"))?;
        if length > MAX_XML_BYTES {
            return Err(invalid("serialized scenarios XML exceeds safety limit"));
        }
        self.bytes
            .try_reserve_exact(value.len())
            .map_err(|_| invalid("serialized scenarios XML output allocation failed"))?;
        self.bytes.extend_from_slice(value);
        Ok(())
    }

    fn finish(self) -> Result<String> {
        String::from_utf8(self.bytes)
            .map_err(|_| invalid("serialized scenarios XML is not valid UTF-8"))
    }
}

fn namespace_kind(result: ResolveResult<'_>) -> Result<NamespaceKind> {
    match result {
        ResolveResult::Unbound => Ok(NamespaceKind::Unbound),
        ResolveResult::Bound(namespace) if is_main_namespace(namespace.as_ref()) => {
            Ok(NamespaceKind::Main)
        },
        ResolveResult::Bound(_) => Ok(NamespaceKind::Other),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix {}",
            String::from_utf8_lossy(&prefix)
        ))),
    }
}

fn is_main_namespace(namespace: &[u8]) -> bool {
    namespace == TRANSITIONAL_MAIN.as_bytes() || namespace == STRICT_MAIN.as_bytes()
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn set_once<T>(slot: &mut Option<T>, value: T, name: &str) -> Result<()> {
    if slot.replace(value).is_some() {
        return Err(invalid(format!("duplicate {name} attribute")));
    }
    Ok(())
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid(format!("{name} must be an XML boolean"))),
    }
}

fn parse_u32(value: &str, name: &str) -> Result<u32> {
    value
        .parse::<u32>()
        .map_err(|_| invalid(format!("{name} must be unsignedInt")))
}

fn parse_sqref(value: &str) -> Result<Vec<RangeReference>> {
    let mut ranges = Vec::new();
    for token in value.split_whitespace() {
        if ranges.len() >= MAX_SQREF_ITEMS {
            return Err(invalid(format!(
                "scenarios sqref exceeds safety limit {MAX_SQREF_ITEMS}"
            )));
        }
        ranges.push(RangeReference::new(token)?);
    }
    if ranges.is_empty() {
        return Err(invalid("scenarios sqref cannot be empty"));
    }
    Ok(ranges)
}
