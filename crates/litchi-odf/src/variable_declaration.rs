//! Bounded, inert semantic inventory of ODF variable declarations.

use crate::datatype::{Boolean, Date, DateTimeOdf, DurationOdf, OdfDurationValue};
use chrono::{DateTime, FixedOffset, NaiveDate};
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_GROUPS: usize = 4_096;
const MAX_DECLARATIONS: usize = 65_536;
const MAX_NAME_BYTES: usize = 65_536;
const MAX_VALUE_BYTES: usize = 1_048_576;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

/// XML part containing a declaration group.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfVariablePart {
    Content,
    Styles,
    Flat,
}

/// Standard body family containing declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfVariableBody {
    Text,
    Spreadsheet,
    Presentation,
    Drawing,
    Chart,
}

/// Header or footer variant containing declarations.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfVariableHeaderFooter {
    Header,
    HeaderFirst,
    HeaderLeft,
    Footer,
    FooterFirst,
    FooterLeft,
}

/// Structural scope in which a declaration group occurs.
#[derive(Clone, Debug, PartialEq, Eq, Hash)]
pub enum OdfVariableScope {
    Body(OdfVariableBody),
    HeaderFooter {
        kind: OdfVariableHeaderFooter,
        master_page_name: Option<String>,
    },
}

/// One of the three variable classes defined by ODF 1.3 section 7.4.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum OdfVariableKind {
    Simple,
    User,
    Sequence,
}

/// Declared ODF value type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum OdfVariableValueType {
    Float,
    Percentage,
    Currency,
    Date,
    Time,
    Boolean,
    String,
    Void,
}

/// Typed date or date-time value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum OdfVariableDateValue {
    Date(NaiveDate),
    DateTime(DateTime<FixedOffset>),
}

/// Typed user-field value retaining its exact lexical representation.
#[derive(Clone, Debug, PartialEq)]
pub enum OdfVariableValue {
    Float { value: f64, lexical: String },
    Percentage { value: f64, lexical: String },
    Currency { value: f64, lexical: String, currency: String },
    Date { value: OdfVariableDateValue, lexical: String },
    Time { value: OdfDurationValue, lexical: String },
    Boolean { value: bool, lexical: String },
    String { value: String },
    Void,
}

impl OdfVariableValue {
    pub fn value_type(&self) -> OdfVariableValueType {
        match self {
            Self::Float { .. } => OdfVariableValueType::Float,
            Self::Percentage { .. } => OdfVariableValueType::Percentage,
            Self::Currency { .. } => OdfVariableValueType::Currency,
            Self::Date { .. } => OdfVariableValueType::Date,
            Self::Time { .. } => OdfVariableValueType::Time,
            Self::Boolean { .. } => OdfVariableValueType::Boolean,
            Self::String { .. } => OdfVariableValueType::String,
            Self::Void => OdfVariableValueType::Void,
        }
    }

    pub fn lexical(&self) -> &str {
        match self {
            Self::Float { lexical, .. }
            | Self::Percentage { lexical, .. }
            | Self::Currency { lexical, .. }
            | Self::Date { lexical, .. }
            | Self::Time { lexical, .. }
            | Self::Boolean { lexical, .. } => lexical,
            Self::String { value } => value,
            Self::Void => "",
        }
    }
}

/// One variable declaration.
#[derive(Clone, Debug, PartialEq)]
pub enum OdfVariableDeclaration {
    Simple {
        name: String,
        value_type: OdfVariableValueType,
    },
    User {
        name: String,
        value: Option<OdfVariableValue>,
        formula: Option<String>,
    },
    Sequence {
        name: String,
        display_outline_level: u8,
        separation_character: Option<char>,
    },
}

impl OdfVariableDeclaration {
    pub fn kind(&self) -> OdfVariableKind {
        match self {
            Self::Simple { .. } => OdfVariableKind::Simple,
            Self::User { .. } => OdfVariableKind::User,
            Self::Sequence { .. } => OdfVariableKind::Sequence,
        }
    }

    pub fn name(&self) -> &str {
        match self {
            Self::Simple { name, .. }
            | Self::User { name, .. }
            | Self::Sequence { name, .. } => name,
        }
    }

    /// Effective separator. Nonzero sequence levels default to `.`.
    pub fn effective_separation_character(&self) -> Option<char> {
        match self {
            Self::Sequence {
                display_outline_level: 1..=10,
                separation_character,
                ..
            } => Some(separation_character.unwrap_or('.')),
            _ => None,
        }
    }
}

/// One declaration container in source order.
#[derive(Clone, Debug, PartialEq)]
pub struct OdfVariableDeclarationGroup {
    pub kind: OdfVariableKind,
    pub part: OdfVariablePart,
    pub scope: OdfVariableScope,
    pub declarations: Vec<OdfVariableDeclaration>,
}

/// Ordered declaration groups from all scanned XML parts.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct OdfVariableDeclarations {
    pub groups: Vec<OdfVariableDeclarationGroup>,
    /// Inert DDE source declarations in document order.
    pub dde_connections: Vec<crate::OdfDdeConnectionDeclaration>,
    /// Validated references to DDE declarations in document order.
    pub dde_connection_uses: Vec<crate::OdfDdeConnectionUse>,
}

impl OdfVariableDeclarations {
    pub fn declarations(&self) -> impl Iterator<Item = &OdfVariableDeclaration> {
        self.groups.iter().flat_map(|group| group.declarations.iter())
    }

    pub fn find(
        &self,
        kind: OdfVariableKind,
        name: &str,
    ) -> Option<&OdfVariableDeclaration> {
        self.declarations()
            .find(|declaration| declaration.kind() == kind && declaration.name() == name)
    }

    pub fn find_dde_connection(&self, name: &str) -> Option<&crate::OdfDdeConnectionDeclaration> {
        self.dde_connections
            .iter()
            .find(|connection| connection.name == name)
    }
}

#[derive(Clone)]
struct Frame {
    namespace: Option<String>,
    local: String,
    master_page_name: Option<String>,
}

struct ActiveGroup {
    depth: usize,
    group: OdfVariableDeclarationGroup,
}

struct PendingDeclaration {
    depth: usize,
    declaration: OdfVariableDeclaration,
}

#[derive(Hash, PartialEq, Eq)]
struct ScopedName {
    part: OdfVariablePart,
    scope: OdfVariableScope,
    kind: OdfVariableKind,
    name: String,
}

pub(crate) fn parse_variable_declaration_parts(
    parts: &[(&str, OdfVariablePart)],
) -> Result<OdfVariableDeclarations> {
    let total = parts.iter().try_fold(0usize, |size, (xml, _)| {
        size.checked_add(xml.len())
            .ok_or_else(|| invalid("variable declaration XML size overflow"))
    })?;
    if total > MAX_XML_BYTES {
        return Err(invalid("variable declaration XML exceeds 64 MiB"));
    }
    let mut result = OdfVariableDeclarations::default();
    let mut names = HashSet::<(OdfVariableKind, String)>::new();
    let mut containers = HashSet::<(OdfVariablePart, OdfVariableScope, OdfVariableKind)>::new();
    let mut uses = HashSet::<ScopedName>::new();
    let mut all_uses = Vec::<(OdfVariableKind, String)>::new();
    let mut aggregate = 0usize;
    let mut declaration_count = 0usize;
    for (xml, part) in parts {
        parse_part(
            xml,
            *part,
            &mut result,
            &mut names,
            &mut containers,
            &mut uses,
            &mut all_uses,
            &mut aggregate,
            &mut declaration_count,
        )?;
    }
    for (kind, name) in all_uses {
        if !names.contains(&(kind, name.clone())) {
            return Err(invalid(format!(
                "ODF {:?} variable '{name}' is used without a declaration",
                kind
            )));
        }
    }
    let dde = crate::dde_connection::parse_dde_connection_parts(parts)?;
    result.dde_connections = dde.declarations;
    result.dde_connection_uses = dde.uses;
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn parse_part(
    xml: &str,
    part: OdfVariablePart,
    result: &mut OdfVariableDeclarations,
    names: &mut HashSet<(OdfVariableKind, String)>,
    containers: &mut HashSet<(OdfVariablePart, OdfVariableScope, OdfVariableKind)>,
    uses: &mut HashSet<ScopedName>,
    all_uses: &mut Vec<(OdfVariableKind, String)>,
    aggregate: &mut usize,
    declaration_count: &mut usize,
) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut stack = Vec::<Frame>::new();
    let mut active: Option<ActiveGroup> = None;
    let mut pending: Option<PendingDeclaration> = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(format!("invalid variable declaration XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if pending.is_some() {
                    return Err(invalid("variable declarations cannot contain elements"));
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_declaration_name(namespace.as_deref(), &local)?;
                if let Some(group) = active.as_mut() {
                    let expected = declaration_local(group.group.kind);
                    if namespace.as_deref() != Some(TEXT)
                        || local != expected
                        || depth != group.depth
                    {
                        return Err(invalid("declaration containers may contain only their declaration elements"));
                    }
                    let declaration = parse_declaration(&reader, element, group.group.kind, aggregate)?;
                    pending = Some(PendingDeclaration {
                        depth: depth + 1,
                        declaration,
                    });
                } else if let Some(kind) = container_kind(namespace.as_deref(), &local) {
                    start_group(
                        &reader,
                        element,
                        part,
                        kind,
                        depth,
                        &stack,
                        result,
                        containers,
                        &mut active,
                    )?;
                } else if let Some(kind) = usage_kind(namespace.as_deref(), &local) {
                    record_use(
                        &reader, element, part, kind, &stack, uses, all_uses, aggregate,
                    )?;
                }
                let master_page_name = if namespace.as_deref() == Some(STYLE)
                    && local == "master-page"
                {
                    optional_attribute(&reader, element, STYLE, "name")?
                } else {
                    None
                };
                stack.push(Frame {
                    namespace,
                    local,
                    master_page_name,
                });
                depth = depth.checked_add(1).ok_or_else(|| invalid("XML depth overflow"))?;
                if depth > MAX_DEPTH {
                    return Err(invalid(format!(
                        "variable declaration XML nesting exceeds {MAX_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                if pending.is_some() {
                    return Err(invalid("variable declarations cannot contain elements"));
                }
                let namespace = namespace_uri(&namespace)?;
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_declaration_name(namespace.as_deref(), &local)?;
                if let Some(group) = active.as_mut() {
                    if namespace.as_deref() != Some(TEXT)
                        || local != declaration_local(group.group.kind)
                        || depth != group.depth
                    {
                        return Err(invalid("declaration containers may contain only their declaration elements"));
                    }
                    let declaration = parse_declaration(&reader, element, group.group.kind, aggregate)?;
                    add_declaration(
                        declaration,
                        &mut group.group,
                        names,
                        uses,
                        part,
                        declaration_count,
                    )?;
                } else if let Some(kind) = container_kind(namespace.as_deref(), &local) {
                    let mut temporary = None;
                    start_group(
                        &reader,
                        element,
                        part,
                        kind,
                        depth,
                        &stack,
                        result,
                        containers,
                        &mut temporary,
                    )?;
                    result.groups.push(temporary.expect("group created").group);
                } else if let Some(kind) = usage_kind(namespace.as_deref(), &local) {
                    record_use(
                        &reader, element, part, kind, &stack, uses, all_uses, aggregate,
                    )?;
                }
            },
            Event::End(_) => {
                if let Some(pending_declaration) = pending.take() {
                    if pending_declaration.depth != depth {
                        pending = Some(pending_declaration);
                    } else {
                        let group = active.as_mut().ok_or_else(|| invalid("orphan declaration"))?;
                        add_declaration(
                            pending_declaration.declaration,
                            &mut group.group,
                            names,
                            uses,
                            part,
                            declaration_count,
                        )?;
                    }
                }
                if active.as_ref().is_some_and(|group| group.depth == depth) {
                    result.groups.push(active.take().expect("checked group").group);
                }
                depth = depth.checked_sub(1).ok_or_else(|| invalid("XML stack underflow"))?;
                stack.pop().ok_or_else(|| invalid("XML frame stack underflow"))?;
            },
            Event::Text(ref text) => {
                let value = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid declaration text: {error}")))?;
                if pending.is_some() && !value.is_empty() {
                    return Err(invalid("declaration elements must have no content"));
                }
                if active.is_some() && pending.is_none() && !value.trim().is_empty() {
                    return Err(invalid("declaration containers may contain only declarations"));
                }
            },
            Event::CData(ref value) if pending.is_some() || active.is_some() => {
                if !value.is_empty() {
                    return Err(invalid("declaration elements cannot contain CDATA"));
                }
            },
            Event::GeneralRef(_) if pending.is_some() || active.is_some() => {
                return Err(invalid("declaration elements cannot contain entity references"));
            },
            Event::DocType(_) => return Err(invalid("DTDs are not allowed in declaration XML")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if depth != 0 || active.is_some() || pending.is_some() || !stack.is_empty() {
        return Err(invalid("incomplete variable declaration XML structure"));
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn start_group(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: OdfVariablePart,
    kind: OdfVariableKind,
    depth: usize,
    stack: &[Frame],
    result: &OdfVariableDeclarations,
    containers: &mut HashSet<(OdfVariablePart, OdfVariableScope, OdfVariableKind)>,
    active: &mut Option<ActiveGroup>,
) -> Result<()> {
    if element.attributes().next().is_some() {
        return Err(invalid("variable declaration containers cannot have attributes"));
    }
    if result.groups.len() >= MAX_GROUPS {
        return Err(invalid(format!("document exceeds {MAX_GROUPS} declaration groups")));
    }
    let parent = stack.last().ok_or_else(|| invalid("misplaced declaration container"))?;
    let scope = scope_for_parent(parent, stack)?;
    if !containers.insert((part, scope.clone(), kind)) {
        return Err(invalid("duplicate variable declaration container in one scope"));
    }
    *active = Some(ActiveGroup {
        depth: depth + 1,
        group: OdfVariableDeclarationGroup {
            kind,
            part,
            scope,
            declarations: Vec::new(),
        },
    });
    let _ = reader;
    Ok(())
}

fn add_declaration(
    declaration: OdfVariableDeclaration,
    group: &mut OdfVariableDeclarationGroup,
    names: &mut HashSet<(OdfVariableKind, String)>,
    uses: &HashSet<ScopedName>,
    part: OdfVariablePart,
    declaration_count: &mut usize,
) -> Result<()> {
    if *declaration_count >= MAX_DECLARATIONS {
        return Err(invalid(format!("document exceeds {MAX_DECLARATIONS} variable declarations")));
    }
    let kind = declaration.kind();
    let name = declaration.name().to_string();
    if uses.contains(&ScopedName {
        part,
        scope: group.scope.clone(),
        kind,
        name: name.clone(),
    }) {
        return Err(invalid(format!(
            "ODF {:?} variable '{name}' is declared after its use",
            kind
        )));
    }
    if !names.insert((kind, name.clone())) {
        return Err(invalid(format!("duplicate ODF {:?} variable declaration '{name}'", kind)));
    }
    *declaration_count += 1;
    group.declarations.push(declaration);
    Ok(())
}

fn parse_declaration(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    kind: OdfVariableKind,
    aggregate: &mut usize,
) -> Result<OdfVariableDeclaration> {
    let attributes = collect_attributes(reader, element, aggregate)?;
    let name = required(&attributes, TEXT, "name")?.to_string();
    validate_string(&name, MAX_NAME_BYTES, "variable name")?;
    match kind {
        OdfVariableKind::Simple => {
            reject_unexpected(&attributes, &[(TEXT, "name"), (OFFICE, "value-type")])?;
            Ok(OdfVariableDeclaration::Simple {
                name,
                value_type: parse_value_type(required(&attributes, OFFICE, "value-type")?)?,
            })
        },
        OdfVariableKind::User => {
            reject_unexpected(
                &attributes,
                &[
                    (TEXT, "name"),
                    (TEXT, "formula"),
                    (OFFICE, "value-type"),
                    (OFFICE, "value"),
                    (OFFICE, "boolean-value"),
                    (OFFICE, "currency"),
                    (OFFICE, "date-value"),
                    (OFFICE, "string-value"),
                    (OFFICE, "time-value"),
                ],
            )?;
            let formula = get(&attributes, TEXT, "formula").map(str::to_string);
            if formula.as_ref().is_some_and(|value| value.len() > MAX_VALUE_BYTES) {
                return Err(invalid("user-field formula exceeds 1 MiB"));
            }
            let value = get(&attributes, OFFICE, "value-type")
                .map(|kind| parse_user_value(parse_value_type(kind)?, &attributes))
                .transpose()?;
            if value.is_none() {
                reject_typed_value_attributes(&attributes, &[])?;
            }
            Ok(OdfVariableDeclaration::User { name, value, formula })
        },
        OdfVariableKind::Sequence => {
            reject_unexpected(
                &attributes,
                &[
                    (TEXT, "name"),
                    (TEXT, "display-outline-level"),
                    (TEXT, "separation-character"),
                ],
            )?;
            let level = required(&attributes, TEXT, "display-outline-level")?
                .parse::<u8>()
                .map_err(|_| invalid("invalid sequence display outline level"))?;
            if level > 10 {
                return Err(invalid("sequence display outline level exceeds 10"));
            }
            let separation_character = get(&attributes, TEXT, "separation-character")
                .map(|value| {
                    let mut chars = value.chars();
                    let character = chars.next().ok_or_else(|| invalid("empty sequence separator"))?;
                    if chars.next().is_some() {
                        return Err(invalid("sequence separator must be one Unicode scalar"));
                    }
                    Ok(character)
                })
                .transpose()?;
            if level == 0 && separation_character.is_some() {
                return Err(invalid("level-zero sequence declaration cannot have a separator"));
            }
            Ok(OdfVariableDeclaration::Sequence {
                name,
                display_outline_level: level,
                separation_character,
            })
        },
    }
}

type Attributes = HashMap<(String, String), String>;

fn collect_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    aggregate: &mut usize,
) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(format!("invalid declaration attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = namespace_uri(&namespace)?.unwrap_or_default();
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| invalid(format!("invalid declaration attribute value: {error}")))?
            .into_owned();
        validate_string(&value, MAX_VALUE_BYTES, "declaration attribute")?;
        *aggregate = aggregate
            .checked_add(value.len())
            .ok_or_else(|| invalid("declaration text size overflow"))?;
        if *aggregate > MAX_AGGREGATE_BYTES {
            return Err(invalid("declaration text exceeds 16 MiB"));
        }
        if result.insert((namespace, local), value).is_some() {
            return Err(invalid("duplicate expanded declaration attribute"));
        }
    }
    Ok(result)
}

fn parse_user_value(kind: OdfVariableValueType, attributes: &Attributes) -> Result<OdfVariableValue> {
    let allowed = match kind {
        OdfVariableValueType::Float | OdfVariableValueType::Percentage => &["value"][..],
        OdfVariableValueType::Currency => &["value", "currency"][..],
        OdfVariableValueType::Date => &["date-value"][..],
        OdfVariableValueType::Time => &["time-value"][..],
        OdfVariableValueType::Boolean => &["boolean-value"][..],
        OdfVariableValueType::String => &["string-value"][..],
        OdfVariableValueType::Void => &[][..],
    };
    reject_typed_value_attributes(attributes, allowed)?;
    let value = match kind {
        OdfVariableValueType::Float => {
            let lexical = required(attributes, OFFICE, "value")?.to_string();
            OdfVariableValue::Float { value: parse_double(&lexical)?, lexical }
        },
        OdfVariableValueType::Percentage => {
            let lexical = required(attributes, OFFICE, "value")?.to_string();
            OdfVariableValue::Percentage { value: parse_double(&lexical)?, lexical }
        },
        OdfVariableValueType::Currency => {
            let lexical = required(attributes, OFFICE, "value")?.to_string();
            let currency = required(attributes, OFFICE, "currency")?.to_string();
            if currency.is_empty() {
                return Err(invalid("currency value requires a currency code"));
            }
            OdfVariableValue::Currency { value: parse_double(&lexical)?, lexical, currency }
        },
        OdfVariableValueType::Date => {
            let lexical = required(attributes, OFFICE, "date-value")?.to_string();
            let value = if lexical.contains('T') {
                OdfVariableDateValue::DateTime(
                    DateTimeOdf::decode(&lexical).map_err(|_| invalid("invalid user-field date-time"))?,
                )
            } else {
                OdfVariableDateValue::Date(
                    Date::decode(&lexical).map_err(|_| invalid("invalid user-field date"))?,
                )
            };
            OdfVariableValue::Date { value, lexical }
        },
        OdfVariableValueType::Time => {
            let lexical = required(attributes, OFFICE, "time-value")?.to_string();
            let value = DurationOdf::decode_exact(&lexical)
                .map_err(|_| invalid("invalid user-field duration"))?;
            OdfVariableValue::Time { value, lexical }
        },
        OdfVariableValueType::Boolean => {
            let lexical = required(attributes, OFFICE, "boolean-value")?.to_string();
            let value = Boolean::decode(&lexical).map_err(|_| invalid("invalid user-field boolean"))?;
            OdfVariableValue::Boolean { value, lexical }
        },
        OdfVariableValueType::String => OdfVariableValue::String {
            value: required(attributes, OFFICE, "string-value")?.to_string(),
        },
        OdfVariableValueType::Void => OdfVariableValue::Void,
    };
    Ok(value)
}

fn reject_typed_value_attributes(attributes: &Attributes, allowed: &[&str]) -> Result<()> {
    for local in [
        "value",
        "boolean-value",
        "currency",
        "date-value",
        "string-value",
        "time-value",
    ] {
        if get(attributes, OFFICE, local).is_some() && !allowed.contains(&local) {
            return Err(invalid(format!(
                "user-field value type does not permit office:{local}"
            )));
        }
    }
    Ok(())
}

fn parse_value_type(value: &str) -> Result<OdfVariableValueType> {
    match value {
        "float" => Ok(OdfVariableValueType::Float),
        "percentage" => Ok(OdfVariableValueType::Percentage),
        "currency" => Ok(OdfVariableValueType::Currency),
        "date" => Ok(OdfVariableValueType::Date),
        "time" => Ok(OdfVariableValueType::Time),
        "boolean" => Ok(OdfVariableValueType::Boolean),
        "string" => Ok(OdfVariableValueType::String),
        "void" => Ok(OdfVariableValueType::Void),
        _ => Err(invalid(format!("unsupported ODF variable value type '{value}'"))),
    }
}

fn parse_double(value: &str) -> Result<f64> {
    match value {
        "INF" => Ok(f64::INFINITY),
        "-INF" => Ok(f64::NEG_INFINITY),
        "NaN" => Ok(f64::NAN),
        _ => value.parse().map_err(|_| invalid("invalid XML Schema double")),
    }
}

fn record_use(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    part: OdfVariablePart,
    kind: OdfVariableKind,
    stack: &[Frame],
    uses: &mut HashSet<ScopedName>,
    all_uses: &mut Vec<(OdfVariableKind, String)>,
    aggregate: &mut usize,
) -> Result<()> {
    let name = required_attribute(reader, element, TEXT, "name")?;
    validate_string(&name, MAX_NAME_BYTES, "variable use name")?;
    *aggregate = aggregate
        .checked_add(name.len())
        .ok_or_else(|| invalid("declaration text size overflow"))?;
    if *aggregate > MAX_AGGREGATE_BYTES {
        return Err(invalid("declaration text exceeds 16 MiB"));
    }
    if let Some(scope) = nearest_scope(stack)? {
        uses.insert(ScopedName {
            part,
            scope,
            kind,
            name: name.clone(),
        });
    }
    all_uses.push((kind, name));
    Ok(())
}

fn scope_for_parent(parent: &Frame, stack: &[Frame]) -> Result<OdfVariableScope> {
    if parent.namespace.as_deref() == Some(OFFICE) {
        let body = match parent.local.as_str() {
            "text" => Some(OdfVariableBody::Text),
            "spreadsheet" => Some(OdfVariableBody::Spreadsheet),
            "presentation" => Some(OdfVariableBody::Presentation),
            "drawing" => Some(OdfVariableBody::Drawing),
            "chart" => Some(OdfVariableBody::Chart),
            _ => None,
        };
        if let Some(body) = body {
            return Ok(OdfVariableScope::Body(body));
        }
    }
    if parent.namespace.as_deref() == Some(STYLE) {
        let kind = match parent.local.as_str() {
            "header" => Some(OdfVariableHeaderFooter::Header),
            "header-first" => Some(OdfVariableHeaderFooter::HeaderFirst),
            "header-left" => Some(OdfVariableHeaderFooter::HeaderLeft),
            "footer" => Some(OdfVariableHeaderFooter::Footer),
            "footer-first" => Some(OdfVariableHeaderFooter::FooterFirst),
            "footer-left" => Some(OdfVariableHeaderFooter::FooterLeft),
            _ => None,
        };
        if let Some(kind) = kind {
            let master_page_name = stack
                .iter()
                .rev()
                .find_map(|frame| frame.master_page_name.clone());
            return Ok(OdfVariableScope::HeaderFooter { kind, master_page_name });
        }
    }
    Err(invalid("misplaced variable declaration container"))
}

fn nearest_scope(stack: &[Frame]) -> Result<Option<OdfVariableScope>> {
    for (index, frame) in stack.iter().enumerate().rev() {
        if frame.namespace.as_deref() == Some(OFFICE)
            && matches!(frame.local.as_str(), "text" | "spreadsheet" | "presentation" | "drawing" | "chart")
        {
            return scope_for_parent(frame, &stack[..=index]).map(Some);
        }
        if frame.namespace.as_deref() == Some(STYLE)
            && matches!(frame.local.as_str(), "header" | "header-first" | "header-left" | "footer" | "footer-first" | "footer-left")
        {
            return scope_for_parent(frame, &stack[..=index]).map(Some);
        }
    }
    Ok(None)
}

fn container_kind(namespace: Option<&str>, local: &str) -> Option<OdfVariableKind> {
    (namespace == Some(TEXT)).then(|| match local {
        "variable-decls" => Some(OdfVariableKind::Simple),
        "user-field-decls" => Some(OdfVariableKind::User),
        "sequence-decls" => Some(OdfVariableKind::Sequence),
        _ => None,
    })?
}

fn declaration_local(kind: OdfVariableKind) -> &'static str {
    match kind {
        OdfVariableKind::Simple => "variable-decl",
        OdfVariableKind::User => "user-field-decl",
        OdfVariableKind::Sequence => "sequence-decl",
    }
}

fn usage_kind(namespace: Option<&str>, local: &str) -> Option<OdfVariableKind> {
    if namespace != Some(TEXT) {
        return None;
    }
    match local {
        "variable-set" | "variable-get" | "variable-input" => Some(OdfVariableKind::Simple),
        "user-field-get" | "user-field-input" => Some(OdfVariableKind::User),
        "sequence" => Some(OdfVariableKind::Sequence),
        _ => None,
    }
}

fn reject_spoofed_declaration_name(namespace: Option<&str>, local: &str) -> Result<()> {
    if matches!(
        local,
        "variable-decls" | "variable-decl" | "user-field-decls" | "user-field-decl" | "sequence-decls" | "sequence-decl"
    ) && namespace != Some(TEXT)
    {
        return Err(invalid("variable declaration vocabulary uses the wrong namespace"));
    }
    Ok(())
}

fn reject_unexpected(attributes: &Attributes, allowed: &[(&str, &str)]) -> Result<()> {
    for (namespace, local) in attributes.keys() {
        if !allowed.iter().any(|(allowed_namespace, allowed_local)| {
            namespace == allowed_namespace && local == allowed_local
        }) && matches!(namespace.as_str(), OFFICE | TEXT)
        {
            return Err(invalid(format!("unexpected declaration attribute {namespace}:{local}")));
        }
    }
    Ok(())
}

fn get<'a>(attributes: &'a Attributes, namespace: &str, local: &str) -> Option<&'a str> {
    attributes
        .get(&(namespace.to_string(), local.to_string()))
        .map(String::as_str)
}

fn required<'a>(attributes: &'a Attributes, namespace: &str, local: &str) -> Result<&'a str> {
    get(attributes, namespace, local)
        .ok_or_else(|| invalid(format!("declaration requires {local}")))
}

fn required_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<String> {
    optional_attribute(reader, element, namespace, local)?
        .ok_or_else(|| invalid(format!("variable use requires {local}")))
}

fn optional_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &str,
    local: &str,
) -> Result<Option<String>> {
    let mut value = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| invalid(format!("invalid XML attribute: {error}")))?;
        let (resolved, resolved_local) = reader.resolver().resolve_attribute(attribute.key);
        if namespace_uri(&resolved)?.as_deref() == Some(namespace)
            && resolved_local.as_ref() == local.as_bytes()
        {
            if value.is_some() {
                return Err(invalid(format!("duplicate expanded attribute {local}")));
            }
            value = Some(
                attribute
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                    .map_err(|error| invalid(format!("invalid XML attribute value: {error}")))?
                    .into_owned(),
            );
        }
    }
    Ok(value)
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(Some(decode(value, "namespace URI")?))
        },
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn decode(value: &[u8], description: &str) -> Result<String> {
    std::str::from_utf8(value)
        .map(str::to_string)
        .map_err(|_| invalid(format!("non-UTF-8 {description}")))
}

fn validate_string(value: &str, limit: usize, description: &str) -> Result<()> {
    if value.is_empty() {
        return Err(invalid(format!("{description} cannot be empty")));
    }
    if value.len() > limit {
        return Err(invalid(format!("{description} exceeds {limit} bytes")));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
