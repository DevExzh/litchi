//! Worksheet what-if scenarios (`CT_Scenarios`, `CT_Scenario`, `CT_InputCells`).
//!
//! A scenario is a named set of substitute values ("input cells") that a user
//! can swap into the model. This module parses and serializes the worksheet
//! `scenarios` collection without evaluating any scenario.

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use crate::error::{Result, invalid};
use litchi_ooxml_common::mce::process_str;

const TRANSITIONAL_MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_MAIN: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const MAX_SCENARIOS: usize = 65_535;
const MAX_INPUT_CELLS: usize = 65_536;
const MAX_SQREF_ITEMS: usize = 32_767;
const MAX_XSTRING_CHARS: usize = 32_767;
const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;
const MAX_ROW: u32 = 1_048_576;
const MAX_COLUMN: u32 = 16_384;

/// Namespace form used when serializing a scenarios fragment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WorksheetScenarioConformance {
    Transitional,
    Strict,
}

impl WorksheetScenarioConformance {
    fn main_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_MAIN,
            Self::Strict => STRICT_MAIN,
        }
    }
}

/// A validated A1 cell reference (`ST_CellRef`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScenarioCellReference(String);

impl ScenarioCellReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_cell_reference(&value, "scenario cell reference")?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// A validated A1 cell or rectangular range reference from `scenarios sqref`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ScenarioRangeReference(String);

impl ScenarioRangeReference {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_range_reference(&value, "scenarios sqref item")?;
        Ok(Self(value))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// One substitute value assignment (`CT_InputCells`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetScenarioInputCell {
    reference: ScenarioCellReference,
    deleted: bool,
    undone: bool,
    value: String,
    number_format_id: Option<u32>,
}

impl WorksheetScenarioInputCell {
    pub fn new(reference: ScenarioCellReference, value: impl Into<String>) -> Result<Self> {
        let value = checked_xstring(value.into(), "inputCells val")?;
        Ok(Self {
            reference,
            deleted: false,
            undone: false,
            value,
            number_format_id: None,
        })
    }

    pub fn with_deleted(mut self, value: bool) -> Self {
        self.deleted = value;
        self
    }
    pub fn with_undone(mut self, value: bool) -> Self {
        self.undone = value;
        self
    }
    pub fn with_number_format_id(mut self, value: u32) -> Self {
        self.number_format_id = Some(value);
        self
    }

    pub fn reference(&self) -> &ScenarioCellReference {
        &self.reference
    }
    pub fn deleted(&self) -> bool {
        self.deleted
    }
    pub fn undone(&self) -> bool {
        self.undone
    }
    pub fn value(&self) -> &str {
        &self.value
    }
    pub fn number_format_id(&self) -> Option<u32> {
        self.number_format_id
    }
}

/// One named what-if scenario (`CT_Scenario`).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetScenario {
    name: String,
    locked: bool,
    hidden: bool,
    count: Option<u32>,
    user: Option<String>,
    comment: Option<String>,
    input_cells: Vec<WorksheetScenarioInputCell>,
}

impl WorksheetScenario {
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let name = checked_xstring(name.into(), "scenario name")?;
        Ok(Self {
            name,
            locked: false,
            hidden: false,
            count: None,
            user: None,
            comment: None,
            input_cells: Vec::new(),
        })
    }

    pub fn with_locked(mut self, value: bool) -> Self {
        self.locked = value;
        self
    }
    pub fn with_hidden(mut self, value: bool) -> Self {
        self.hidden = value;
        self
    }
    pub fn with_count(mut self, value: u32) -> Self {
        self.count = Some(value);
        self
    }
    pub fn with_user(mut self, value: impl Into<String>) -> Result<Self> {
        self.user = Some(checked_xstring(value.into(), "scenario user")?);
        Ok(self)
    }
    pub fn with_comment(mut self, value: impl Into<String>) -> Result<Self> {
        self.comment = Some(checked_xstring(value.into(), "scenario comment")?);
        Ok(self)
    }
    pub fn with_input_cells(mut self, value: Vec<WorksheetScenarioInputCell>) -> Result<Self> {
        if value.len() > MAX_INPUT_CELLS {
            return Err(invalid(format!(
                "scenario inputCells exceeds safety limit {MAX_INPUT_CELLS}"
            )));
        }
        self.input_cells = value;
        Ok(self)
    }

    pub fn name(&self) -> &str {
        &self.name
    }
    pub fn locked(&self) -> bool {
        self.locked
    }
    pub fn hidden(&self) -> bool {
        self.hidden
    }
    pub fn count(&self) -> Option<u32> {
        self.count
    }
    pub fn user(&self) -> Option<&str> {
        self.user.as_deref()
    }
    pub fn comment(&self) -> Option<&str> {
        self.comment.as_deref()
    }
    pub fn input_cells(&self) -> &[WorksheetScenarioInputCell] {
        &self.input_cells
    }
}

/// The worksheet `scenarios` collection in document order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WorksheetScenarios {
    current: Option<u32>,
    show: Option<u32>,
    ranges: Vec<ScenarioRangeReference>,
    scenarios: Vec<WorksheetScenario>,
}

impl WorksheetScenarios {
    pub fn new(scenarios: Vec<WorksheetScenario>) -> Result<Self> {
        if scenarios.is_empty() {
            return Err(invalid("scenarios requires at least one scenario"));
        }
        if scenarios.len() > MAX_SCENARIOS {
            return Err(invalid(format!(
                "scenarios exceeds safety limit {MAX_SCENARIOS}"
            )));
        }
        Ok(Self {
            current: None,
            show: None,
            ranges: Vec::new(),
            scenarios,
        })
    }

    pub fn with_current(mut self, value: u32) -> Self {
        self.current = Some(value);
        self
    }
    pub fn with_show(mut self, value: u32) -> Self {
        self.show = Some(value);
        self
    }
    pub fn with_ranges(mut self, value: Vec<ScenarioRangeReference>) -> Result<Self> {
        if value.len() > MAX_SQREF_ITEMS {
            return Err(invalid(format!(
                "scenarios sqref exceeds safety limit {MAX_SQREF_ITEMS}"
            )));
        }
        self.ranges = value;
        Ok(self)
    }

    pub fn current(&self) -> Option<u32> {
        self.current
    }
    pub fn show(&self) -> Option<u32> {
        self.show
    }
    pub fn ranges(&self) -> &[ScenarioRangeReference] {
        &self.ranges
    }
    pub fn scenarios(&self) -> &[WorksheetScenario] {
        &self.scenarios
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scope {
    Worksheet,
    Scenarios,
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

/// Parses the direct worksheet `scenarios` child after applying shared MCE processing.
pub fn parse_worksheet_scenarios(xml: &[u8]) -> Result<Option<WorksheetScenarios>> {
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
    let mut state: Option<ScenariosBuilder> = None;
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
                    Scope::Scenarios => {
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
                    Some(Scope::Scenarios | Scope::Scenario | Scope::InputCells)
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
                    Some(Scope::Scenarios | Scope::Scenario | Scope::InputCells)
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
struct ScenariosBuilder {
    current: Option<u32>,
    show: Option<u32>,
    ranges: Vec<ScenarioRangeReference>,
    scenarios: Vec<WorksheetScenario>,
    open_scenario: Option<ScenarioBuilder>,
}

#[derive(Default)]
struct ScenarioBuilder {
    name: Option<String>,
    locked: Option<bool>,
    hidden: Option<bool>,
    count: Option<u32>,
    user: Option<String>,
    comment: Option<String>,
    input_cells: Vec<WorksheetScenarioInputCell>,
}

fn begin_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: NamespaceKind,
    parent: Option<Scope>,
    state: &mut Option<ScenariosBuilder>,
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
            Ok(Scope::Scenarios)
        },
        Some(Scope::Scenarios) => {
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
                builder
                    .input_cells
                    .push(parse_input_cell_attributes(reader, element)?);
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

fn end_scope(scope: Scope, state: &mut Option<ScenariosBuilder>) -> Result<()> {
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
    builder.scenarios.push(WorksheetScenario {
        name,
        locked: scenario.locked.unwrap_or(false),
        hidden: scenario.hidden.unwrap_or(false),
        count: scenario.count,
        user: scenario.user,
        comment: scenario.comment,
        input_cells: scenario.input_cells,
    });
    Ok(())
}

fn finish_builder(builder: ScenariosBuilder) -> Result<WorksheetScenarios> {
    if builder.scenarios.is_empty() {
        return Err(invalid("scenarios requires at least one scenario"));
    }
    Ok(WorksheetScenarios {
        current: builder.current,
        show: builder.show,
        ranges: builder.ranges,
        scenarios: builder.scenarios,
    })
}

fn parse_scenarios_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<ScenariosBuilder> {
    let mut value = ScenariosBuilder::default();
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid scenarios attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound {
            return Err(invalid("unknown namespaced scenarios attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid scenarios attribute value: {error}")))?;
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
            _ => return Err(invalid("unknown scenarios attribute")),
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
        if namespace_kind(namespace)? != NamespaceKind::Unbound {
            return Err(invalid("unknown namespaced scenario attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid scenario attribute value: {error}")))?
            .into_owned();
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
            _ => return Err(invalid("unknown scenario attribute")),
        }
    }
    Ok(value)
}

fn parse_input_cell_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
) -> Result<WorksheetScenarioInputCell> {
    let mut reference = None;
    let mut deleted = None;
    let mut undone = None;
    let mut input_value = None;
    let mut number_format_id = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("invalid inputCells attribute: {error}")))?;
        if is_namespace_declaration(attribute.key.as_ref()) {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_kind(namespace)? != NamespaceKind::Unbound {
            return Err(invalid("unknown namespaced inputCells attribute"));
        }
        let text = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid inputCells attribute value: {error}")))?
            .into_owned();
        match local.as_ref() {
            b"r" => set_once(&mut reference, ScenarioCellReference::new(text)?, "r")?,
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
            _ => return Err(invalid("unknown inputCells attribute")),
        }
    }
    Ok(WorksheetScenarioInputCell {
        reference: reference.ok_or_else(|| invalid("inputCells requires r"))?,
        deleted: deleted.unwrap_or(false),
        undone: undone.unwrap_or(false),
        value: input_value.ok_or_else(|| invalid("inputCells requires val"))?,
        number_format_id,
    })
}

/// Serializes one canonical, namespace-complete `scenarios` fragment.
pub fn write_worksheet_scenarios(
    value: &WorksheetScenarios,
    conformance: WorksheetScenarioConformance,
) -> Result<String> {
    if value.scenarios.is_empty() {
        return Err(invalid("scenarios requires at least one scenario"));
    }
    if value.scenarios.len() > MAX_SCENARIOS {
        return Err(invalid(format!(
            "scenarios exceeds safety limit {MAX_SCENARIOS}"
        )));
    }
    let mut xml = String::new();
    xml.push_str("<scenarios xmlns=\"");
    xml.push_str(conformance.main_namespace());
    xml.push('"');
    if let Some(current) = value.current {
        write_u32_attribute(&mut xml, "current", current);
    }
    if let Some(show) = value.show {
        write_u32_attribute(&mut xml, "show", show);
    }
    if !value.ranges.is_empty() {
        xml.push_str(" sqref=\"");
        for (index, range) in value.ranges.iter().enumerate() {
            if index > 0 {
                xml.push(' ');
            }
            xml.push_str(range.as_str());
        }
        xml.push('"');
    }
    xml.push('>');
    for scenario in &value.scenarios {
        xml.push_str("<scenario");
        write_attribute(&mut xml, "name", &scenario.name);
        write_true_attribute(&mut xml, "locked", scenario.locked);
        write_true_attribute(&mut xml, "hidden", scenario.hidden);
        if let Some(count) = scenario.count {
            write_u32_attribute(&mut xml, "count", count);
        }
        if let Some(user) = &scenario.user {
            write_attribute(&mut xml, "user", user);
        }
        if let Some(comment) = &scenario.comment {
            write_attribute(&mut xml, "comment", comment);
        }
        if scenario.input_cells.is_empty() {
            xml.push_str("/>");
            continue;
        }
        xml.push('>');
        for cell in &scenario.input_cells {
            xml.push_str("<inputCells");
            write_attribute(&mut xml, "r", cell.reference.as_str());
            write_true_attribute(&mut xml, "deleted", cell.deleted);
            write_true_attribute(&mut xml, "undone", cell.undone);
            write_attribute(&mut xml, "val", &cell.value);
            if let Some(number_format_id) = cell.number_format_id {
                write_u32_attribute(&mut xml, "numFmtId", number_format_id);
            }
            xml.push_str("/>");
        }
        xml.push_str("</scenario>");
    }
    xml.push_str("</scenarios>");
    Ok(xml)
}

fn write_true_attribute(xml: &mut String, name: &str, value: bool) {
    if value {
        xml.push(' ');
        xml.push_str(name);
        xml.push_str("=\"1\"");
    }
}

fn write_u32_attribute(xml: &mut String, name: &str, value: u32) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&value.to_string());
    xml.push('"');
}

fn write_attribute(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    for character in value.chars() {
        match character {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '>' => xml.push_str("&gt;"),
            '"' => xml.push_str("&quot;"),
            '\'' => xml.push_str("&apos;"),
            _ => xml.push(character),
        }
    }
    xml.push('"');
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

fn checked_xstring(value: String, name: &str) -> Result<String> {
    if value.chars().count() > MAX_XSTRING_CHARS {
        return Err(invalid(format!(
            "{name} exceeds {MAX_XSTRING_CHARS} characters"
        )));
    }
    Ok(value)
}

fn parse_sqref(value: &str) -> Result<Vec<ScenarioRangeReference>> {
    let mut ranges = Vec::new();
    for token in value.split_whitespace() {
        if ranges.len() >= MAX_SQREF_ITEMS {
            return Err(invalid(format!(
                "scenarios sqref exceeds safety limit {MAX_SQREF_ITEMS}"
            )));
        }
        ranges.push(ScenarioRangeReference::new(token)?);
    }
    if ranges.is_empty() {
        return Err(invalid("scenarios sqref cannot be empty"));
    }
    Ok(ranges)
}

fn validate_range_reference(value: &str, name: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if parts.next().is_some() || first.is_empty() || second.is_some_and(str::is_empty) {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    validate_cell_reference(first, name)?;
    if let Some(second) = second {
        validate_cell_reference(second, name)?;
    }
    Ok(())
}

fn validate_cell_reference(value: &str, name: &str) -> Result<()> {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'$'));
    let column_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_alphabetic) {
        index += 1;
    }
    if index == column_start || index - column_start > 3 {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    let mut column = 0u32;
    for byte in &bytes[column_start..index] {
        column = column * 26 + u32::from(byte.to_ascii_uppercase() - b'A' + 1);
    }
    if column == 0 || column > MAX_COLUMN {
        return Err(invalid(format!(
            "{name} column is out of range in '{value}'"
        )));
    }
    if bytes.get(index) == Some(&b'$') {
        index += 1;
    }
    let row_start = index;
    while bytes.get(index).is_some_and(u8::is_ascii_digit) {
        index += 1;
    }
    if index == row_start || index != bytes.len() {
        return Err(invalid(format!("invalid {name} '{value}'")));
    }
    let row = value[row_start..]
        .parse::<u32>()
        .map_err(|_| invalid(format!("invalid {name} row in '{value}'")))?;
    if row == 0 || row > MAX_ROW {
        return Err(invalid(format!("{name} row is out of range in '{value}'")));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    fn parse(child: &str) -> Result<Option<WorksheetScenarios>> {
        parse_worksheet_scenarios(
            format!(r#"<worksheet xmlns="{NS}">{child}</worksheet>"#).as_bytes(),
        )
    }

    #[test]
    fn parses_scenarios_attributes_input_cells_and_defaults() {
        let value = parse(concat!(
            r#"<scenarios current="1" show="1" sqref="A1 $B$2:C3">"#,
            r#"<scenario name="first"><inputCells r="A1" val="10"/></scenario>"#,
            r#"<scenario name="second" locked="1" hidden="true" count="2" user="one" comment="note">"#,
            r#"<inputCells r="B2" val="x" numFmtId="14"/>"#,
            r#"<inputCells r="$C$3" deleted="1" undone="true" val="y &amp; z"/></scenario>"#,
            r#"</scenarios>"#,
        ))
        .unwrap()
        .unwrap();
        assert_eq!(value.current(), Some(1));
        assert_eq!(value.show(), Some(1));
        assert_eq!(
            value
                .ranges()
                .iter()
                .map(ScenarioRangeReference::as_str)
                .collect::<Vec<_>>(),
            vec!["A1", "$B$2:C3"]
        );
        assert_eq!(value.scenarios().len(), 2);
        let first = &value.scenarios()[0];
        assert_eq!(first.name(), "first");
        assert!(!first.locked());
        assert!(!first.hidden());
        assert_eq!(first.count(), None);
        assert_eq!(first.user(), None);
        assert_eq!(first.input_cells()[0].reference().as_str(), "A1");
        assert_eq!(first.input_cells()[0].value(), "10");
        let second = &value.scenarios()[1];
        assert!(second.locked());
        assert!(second.hidden());
        assert_eq!(second.count(), Some(2));
        assert_eq!(second.user(), Some("one"));
        assert_eq!(second.comment(), Some("note"));
        let cells = second.input_cells();
        assert_eq!(cells[0].number_format_id(), Some(14));
        assert!(cells[1].deleted());
        assert!(cells[1].undone());
        assert_eq!(cells[1].value(), "y & z");
    }

    #[test]
    fn supports_strict_namespace_and_skips_extension_markup() {
        let xml = concat!(
            r#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">"#,
            r#"<scenarios><scenario name="s">"#,
            r#"<extLst><ext uri="urn:test"><x:payload xmlns:x="urn:x"/></ext></extLst>"#,
            r#"</scenario></scenarios></worksheet>"#,
        );
        let value = parse_worksheet_scenarios(xml.as_bytes()).unwrap().unwrap();
        assert_eq!(value.scenarios()[0].name(), "s");
    }

    #[test]
    fn rejects_structure_attributes_and_limits() {
        for child in [
            "<scenarios/>",
            r#"<scenarios><scenario/></scenarios>"#,
            r#"<scenarios><scenario name=""><inputCells val="1"/></scenario></scenarios>"#,
            r#"<scenarios><scenario name="s"><inputCells r="A1"/></scenario></scenarios>"#,
            r#"<scenarios><scenario name="s"><inputCells r="A0" val="1"/></scenario></scenarios>"#,
            r#"<scenarios><scenario name="s"><inputCells r="XFE1" val="1"/></scenario></scenarios>"#,
            r#"<scenarios><scenario name="s"><inputCells r="A1" val="1"><child/></inputCells></scenario></scenarios>"#,
            r#"<scenarios><scenario name="s"><inputCells r="A1" val="1" mystery="1"/></scenario></scenarios>"#,
            r#"<scenarios current="yes"><scenario name="s"/></scenarios>"#,
            r#"<scenarios sqref=""><scenario name="s"/></scenarios>"#,
            r#"<scenarios><scenario name="s" locked="yes"/></scenarios>"#,
            r#"<scenarios><scenario name="s">text</scenario></scenarios>"#,
        ] {
            assert!(parse(child).is_err(), "expected rejection for {child}");
        }
        assert!(parse("<scenarios><scenario name=\"s\"/></scenarios><scenarios><scenario name=\"t\"/></scenarios>").is_err());
        let long_name = "x".repeat(MAX_XSTRING_CHARS + 1);
        assert!(
            parse(&format!(
                r#"<scenarios><scenario name="{long_name}"/></scenarios>"#
            ))
            .is_err()
        );
    }

    #[test]
    fn rejects_malformed_document_boundaries_and_excessive_depth() {
        for xml in [
            format!(r#"<worksheet xmlns="{NS}"/><worksheet xmlns="{NS}"/>"#),
            format!(r#"text<worksheet xmlns="{NS}"></worksheet>"#),
            format!(r#"<worksheet xmlns="{NS}">text</worksheet>"#),
            format!(r#"<worksheet xmlns="{NS}"></worksheet>tail"#),
            format!(r#"<worksheet xmlns="{NS}"><![CDATA[data]]></worksheet>"#),
            format!(
                r#"<worksheet xmlns="{NS}"><scenarios><scenario name="s"/></scenarios></worksheet><?pi?>"#
            ),
        ] {
            assert!(
                parse_worksheet_scenarios(xml.as_bytes()).is_err(),
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
        assert!(parse_worksheet_scenarios(xml.as_bytes()).is_err());
    }

    #[test]
    fn write_round_trips_through_the_reader() {
        let scenario = WorksheetScenario::new("baseline")
            .unwrap()
            .with_locked(true)
            .with_count(1)
            .with_user("analyst")
            .unwrap()
            .with_comment("Q1 <plan> & \"notes\"")
            .unwrap()
            .with_input_cells(vec![
                WorksheetScenarioInputCell::new(ScenarioCellReference::new("A1").unwrap(), "10")
                    .unwrap()
                    .with_number_format_id(14),
                WorksheetScenarioInputCell::new(
                    ScenarioCellReference::new("$B$2").unwrap(),
                    "hold",
                )
                .unwrap()
                .with_deleted(true)
                .with_undone(true),
            ])
            .unwrap();
        let expected = WorksheetScenarios::new(vec![scenario])
            .unwrap()
            .with_current(0)
            .with_show(0)
            .with_ranges(vec![ScenarioRangeReference::new("A1:B2").unwrap()])
            .unwrap();
        for conformance in [
            WorksheetScenarioConformance::Transitional,
            WorksheetScenarioConformance::Strict,
        ] {
            let fragment = write_worksheet_scenarios(&expected, conformance).unwrap();
            let document = format!(r#"<worksheet xmlns="{NS}">{fragment}</worksheet>"#);
            let parsed = parse_worksheet_scenarios(document.as_bytes())
                .unwrap()
                .unwrap();
            assert_eq!(parsed, expected);
        }
    }

    #[test]
    fn reads_libreoffice_scenario_fixture() {
        let package = OpcPackage::from_bytes(include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf97598_scenarios.xlsx"
        )))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        let value = parse_worksheet_scenarios(part.blob()).unwrap().unwrap();
        assert_eq!(value.current(), Some(0));
        assert_eq!(value.scenarios().len(), 1);
        let scenario = &value.scenarios()[0];
        assert_eq!(scenario.name(), "scenario1");
        assert!(scenario.locked());
        assert_eq!(scenario.count(), Some(1));
        assert_eq!(scenario.user(), Some("one"));
        assert_eq!(scenario.comment(), Some("Created by one on 12/26/2016"));
        assert_eq!(scenario.input_cells().len(), 1);
        assert_eq!(scenario.input_cells()[0].reference().as_str(), "A1");
        assert_eq!(scenario.input_cells()[0].value(), "value from scenario1");
        // Sheets without scenarios parse to None.
        let other = package
            .get_part(&PackURI::new("/xl/worksheets/sheet2.xml").unwrap())
            .unwrap();
        assert!(parse_worksheet_scenarios(other.blob()).unwrap().is_none());
    }
}
