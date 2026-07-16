//! Typed, inert ODF spreadsheet named ranges and expressions.

use litchi_core::{Error, Result};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use quick_xml::XmlVersion;
use std::collections::{HashMap, HashSet};
use std::io::BufRead;

const OFFICE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NS: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_XML_BYTES: usize = 64 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_GROUPS: usize = 65_536;
const MAX_DEFINITIONS: usize = 262_144;
const MAX_VALUE_BYTES: usize = 65_536;
const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfNamedExpressionPart { Content, FlatDocument }

#[derive(Clone, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfNamedExpressionScope { Spreadsheet, Table { name: Option<String> } }

macro_rules! lexical_type {
    ($name:ident, $label:literal, $empty:expr) => {
        #[derive(Clone, Debug, PartialEq, Eq, Hash)]
        pub struct $name(String);
        impl $name {
            pub fn new(value: impl Into<String>) -> Result<Self> {
                let value = value.into();
                validate_text(&value, $label, $empty)?;
                Ok(Self(value))
            }
            pub fn as_str(&self) -> &str { &self.0 }
        }
    };
}

lexical_type!(OdfCellAddress, "table:base-cell-address", false);
lexical_type!(OdfCellRangeAddress, "table:cell-range-address", false);
lexical_type!(OdfFormulaExpression, "table:expression", true);

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OdfNamedRangeUse { PrintRange, Filter, RepeatRow, RepeatColumn }

impl OdfNamedRangeUse {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "print-range" => Ok(Self::PrintRange),
            "filter" => Ok(Self::Filter),
            "repeat-row" => Ok(Self::RepeatRow),
            "repeat-column" => Ok(Self::RepeatColumn),
            _ => invalid(format!("unsupported table:range-usable-as token '{value}'")),
        }
    }
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::PrintRange => "print-range",
            Self::Filter => "filter",
            Self::RepeatRow => "repeat-row",
            Self::RepeatColumn => "repeat-column",
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfNamedRangeUsage { None, Uses(Vec<OdfNamedRangeUse>) }

impl OdfNamedRangeUsage {
    pub fn uses(values: Vec<OdfNamedRangeUse>) -> Result<Self> {
        if values.is_empty() { return invalid("named range usage list cannot be empty"); }
        let mut unique = HashSet::with_capacity(values.len());
        if values.iter().any(|value| !unique.insert(*value)) {
            return invalid("named range usage tokens must be unique");
        }
        Ok(Self::Uses(values))
    }
    fn parse(value: &str) -> Result<Self> {
        if value == "none" { return Ok(Self::None); }
        Self::uses(value.split_ascii_whitespace().map(OdfNamedRangeUse::parse).collect::<Result<Vec<_>>>()?)
    }
    fn validate(&self) -> Result<()> {
        match self { Self::None => Ok(()), Self::Uses(values) => Self::uses(values.clone()).map(|_| ()) }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OdfNamedDefinition {
    Range {
        name: String,
        cell_range_address: OdfCellRangeAddress,
        base_cell_address: Option<OdfCellAddress>,
        usage: Option<OdfNamedRangeUsage>,
    },
    Expression {
        name: String,
        expression: OdfFormulaExpression,
        base_cell_address: Option<OdfCellAddress>,
    },
}

impl OdfNamedDefinition {
    pub fn name(&self) -> &str {
        match self { Self::Range { name, .. } | Self::Expression { name, .. } => name }
    }
    fn validate(&self) -> Result<()> {
        validate_text(self.name(), "table:name", false)?;
        match self {
            Self::Range { cell_range_address, base_cell_address, usage, .. } => {
                validate_text(cell_range_address.as_str(), "table:cell-range-address", false)?;
                if let Some(value) = base_cell_address { validate_text(value.as_str(), "table:base-cell-address", false)?; }
                if let Some(value) = usage { value.validate()?; }
            }
            Self::Expression { expression, base_cell_address, .. } => {
                validate_text(expression.as_str(), "table:expression", true)?;
                if let Some(value) = base_cell_address { validate_text(value.as_str(), "table:base-cell-address", false)?; }
            }
        }
        Ok(())
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct OdfNamedExpressionGroup {
    pub part: OdfNamedExpressionPart,
    pub scope: OdfNamedExpressionScope,
    pub definitions: Vec<OdfNamedDefinition>,
}

impl OdfNamedExpressionGroup {
    pub fn get(&self, name: &str) -> Option<&OdfNamedDefinition> {
        self.definitions.iter().find(|definition| definition.name() == name)
    }
    pub fn validate(&self) -> Result<()> {
        if self.definitions.len() > MAX_DEFINITIONS { return invalid(format!("named-expression group exceeds {MAX_DEFINITIONS} definitions")); }
        if let OdfNamedExpressionScope::Table { name: Some(name) } = &self.scope { validate_text(name, "table:table table:name", false)?; }
        let mut names = HashSet::with_capacity(self.definitions.len());
        for definition in &self.definitions {
            definition.validate()?;
            if !names.insert(definition.name()) { return invalid(format!("duplicate named definition '{}' in one scope", definition.name())); }
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut output = String::with_capacity(192 + self.definitions.len() * 160);
        output.push_str(r#"<table:named-expressions xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">"#);
        for definition in &self.definitions {
            match definition {
                OdfNamedDefinition::Range { name, cell_range_address, base_cell_address, usage } => {
                    output.push_str("<table:named-range table:name=\"");
                    escape_attribute(&mut output, name);
                    output.push_str("\" table:cell-range-address=\"");
                    escape_attribute(&mut output, cell_range_address.as_str());
                    output.push('"');
                    if let Some(address) = base_cell_address {
                        output.push_str(" table:base-cell-address=\"");
                        escape_attribute(&mut output, address.as_str());
                        output.push('"');
                    }
                    if let Some(usage) = usage {
                        output.push_str(" table:range-usable-as=\"");
                        match usage {
                            OdfNamedRangeUsage::None => output.push_str("none"),
                            OdfNamedRangeUsage::Uses(values) => for (index, value) in values.iter().enumerate() {
                                if index != 0 { output.push(' '); }
                                output.push_str(value.as_str());
                            },
                        }
                        output.push('"');
                    }
                    output.push_str("/>");
                }
                OdfNamedDefinition::Expression { name, expression, base_cell_address } => {
                    output.push_str("<table:named-expression table:name=\"");
                    escape_attribute(&mut output, name);
                    output.push_str("\" table:expression=\"");
                    escape_attribute(&mut output, expression.as_str());
                    output.push('"');
                    if let Some(address) = base_cell_address {
                        output.push_str(" table:base-cell-address=\"");
                        escape_attribute(&mut output, address.as_str());
                        output.push('"');
                    }
                    output.push_str("/>");
                }
            }
        }
        output.push_str("</table:named-expressions>");
        Ok(output)
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OdfNamedExpressions { pub groups: Vec<OdfNamedExpressionGroup> }

impl OdfNamedExpressions {
    pub fn validate(&self) -> Result<()> {
        if self.groups.len() > MAX_GROUPS { return invalid(format!("document exceeds {MAX_GROUPS} named-expression groups")); }
        let mut count = 0usize;
        let mut aggregate = 0usize;
        for group in &self.groups {
            group.validate()?;
            count = count.checked_add(group.definitions.len()).ok_or_else(|| make_error("named definition count overflow"))?;
            if count > MAX_DEFINITIONS { return invalid(format!("document exceeds {MAX_DEFINITIONS} named definitions")); }
            if let OdfNamedExpressionScope::Table { name: Some(name) } = &group.scope {
                aggregate = aggregate.checked_add(name.len()).ok_or_else(|| make_error("named definition text size overflow"))?;
            }
            for definition in &group.definitions {
                aggregate = aggregate.checked_add(definition.name().len()).ok_or_else(|| make_error("named definition text size overflow"))?;
                let extra = match definition {
                    OdfNamedDefinition::Range { cell_range_address, base_cell_address, .. } => cell_range_address.as_str().len() + base_cell_address.as_ref().map_or(0, |value| value.as_str().len()),
                    OdfNamedDefinition::Expression { expression, base_cell_address, .. } => expression.as_str().len() + base_cell_address.as_ref().map_or(0, |value| value.as_str().len()),
                };
                aggregate = aggregate.checked_add(extra).ok_or_else(|| make_error("named definition text size overflow"))?;
            }
            if aggregate > MAX_AGGREGATE_BYTES { return invalid("named definition text exceeds 16 MiB"); }
        }
        Ok(())
    }
}

impl crate::OpenDocumentPackage {
    pub fn named_expressions(&self) -> Result<OdfNamedExpressions> {
        parse_part(&self.content_xml()?, OdfNamedExpressionPart::Content)
    }
}

impl crate::FlatOpenDocument {
    pub fn named_expressions(&self) -> Result<OdfNamedExpressions> { parse_named_expressions(self.xml()) }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
enum NamespaceKind { None, Office, Table, Other }
#[derive(Clone, Debug)]
enum ScopeFrame { Spreadsheet(usize), Table(usize, Option<String>) }
#[derive(Clone, Debug)]
struct Frame { namespace: NamespaceKind, local: String, scope: Option<ScopeFrame> }
struct ActiveGroup { parent_depth: usize, scope_id: usize, value: OdfNamedExpressionGroup }
struct ActiveDefinition { parent_depth: usize, value: OdfNamedDefinition }
type Attributes = HashMap<(NamespaceKind, String), String>;

pub fn parse_named_expressions(xml: &str) -> Result<OdfNamedExpressions> {
    parse_part(xml, OdfNamedExpressionPart::FlatDocument)
}

fn parse_part(xml: &str, part: OdfNamedExpressionPart) -> Result<OdfNamedExpressions> {
    if xml.len() > MAX_XML_BYTES { return invalid("named-expression XML exceeds 64 MiB"); }
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut group: Option<ActiveGroup> = None;
    let mut definition: Option<ActiveDefinition> = None;
    let mut seen_scopes = HashSet::new();
    let mut next_scope_id = 0usize;
    let mut result = OdfNamedExpressions::default();
    loop {
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| make_error(format!("invalid named-expression XML: {error}")))?;
        let namespace = namespace_kind(&resolved)?;
        match event {
            Event::Start(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let needed = group.is_some() || (namespace == NamespaceKind::Table && matches!(local.as_str(), "table" | "named-expressions" | "named-range" | "named-expression"));
                let attributes = if needed { read_attributes(&mut reader, element)? } else { Attributes::new() };
                let scope = handle_start(namespace, &local, attributes, part, stack.len(), stack.last(), &mut group, &mut definition, &seen_scopes, &mut next_scope_id)?;
                stack.push(Frame { namespace, local, scope });
                if stack.len() > MAX_DEPTH { return invalid(format!("named-expression XML exceeds depth {MAX_DEPTH}")); }
            }
            Event::Empty(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                reject_spoofed_name(namespace, &local)?;
                let needed = group.is_some() || (namespace == NamespaceKind::Table && matches!(local.as_str(), "named-expressions" | "named-range" | "named-expression"));
                let attributes = if needed { read_attributes(&mut reader, element)? } else { Attributes::new() };
                handle_empty(namespace, &local, attributes, part, stack.len(), stack.last(), &mut group, &mut definition, &mut seen_scopes, &mut result)?;
            }
            Event::End(ref element) => {
                let local = decode(element.local_name().as_ref(), "element name")?;
                let frame = stack.pop().ok_or_else(|| make_error("unexpected named-expression end element"))?;
                if frame.namespace != namespace || frame.local != local { return invalid("named-expression end element mismatch"); }
                if definition.as_ref().is_some_and(|value| value.parent_depth == stack.len()) {
                    let value = definition.take().expect("checked").value;
                    push_definition(group.as_mut().ok_or_else(|| make_error("named definition has no group"))?, value)?;
                }
                if group.as_ref().is_some_and(|value| value.parent_depth == stack.len()) {
                    let value = group.take().expect("checked");
                    seen_scopes.insert(value.scope_id);
                    result.groups.push(value.value);
                    if result.groups.len() > MAX_GROUPS { return invalid(format!("document exceeds {MAX_GROUPS} named-expression groups")); }
                }
            }
            Event::Text(ref text) => if definition.is_some() || group.is_some() {
                let value = text.decode().map_err(|error| make_error(format!("invalid named-expression text: {error}")))?;
                if !value.trim().is_empty() { return invalid("named definition elements must have empty content"); }
            },
            Event::CData(_) if definition.is_some() || group.is_some() => return invalid("CDATA is not allowed in named definitions"),
            Event::GeneralRef(_) if definition.is_some() || group.is_some() => return invalid("entity references are not allowed in named definitions"),
            Event::DocType(_) => return invalid("DTDs are not allowed in named-expression XML"),
            Event::PI(_) => return invalid("processing instructions are not allowed in named-expression XML"),
            Event::Eof => break,
            _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() || group.is_some() || definition.is_some() { return invalid("truncated named-expression XML"); }
    result.validate()?;
    Ok(result)
}

#[allow(clippy::too_many_arguments)]
fn handle_start(namespace: NamespaceKind, local: &str, attributes: Attributes, part: OdfNamedExpressionPart, depth: usize, parent: Option<&Frame>, group: &mut Option<ActiveGroup>, definition: &mut Option<ActiveDefinition>, seen: &HashSet<usize>, next_id: &mut usize) -> Result<Option<ScopeFrame>> {
    if definition.is_some() { return invalid("named range and expression elements must have empty content"); }
    if let Some(active) = group.as_mut() {
        if depth != active.parent_depth + 1 { return invalid("named definitions must be direct children of table:named-expressions"); }
        *definition = Some(ActiveDefinition { parent_depth: depth, value: parse_definition(namespace, local, attributes)? });
        return Ok(None);
    }
    if namespace == NamespaceKind::Table && local == "named-expressions" {
        if !attributes.is_empty() { return invalid("table:named-expressions does not allow attributes"); }
        let (scope_id, scope) = named_scope(parent)?;
        if seen.contains(&scope_id) { return invalid("a spreadsheet or table may contain only one table:named-expressions group"); }
        *group = Some(ActiveGroup { parent_depth: depth, scope_id, value: OdfNamedExpressionGroup { part, scope, definitions: Vec::new() } });
        return Ok(None);
    }
    if namespace == NamespaceKind::Table && matches!(local, "named-range" | "named-expression") { return invalid("named definitions must be direct children of table:named-expressions"); }
    let scope = if namespace == NamespaceKind::Office && local == "spreadsheet" {
        Some(ScopeFrame::Spreadsheet(allocate_id(next_id)?))
    } else if namespace == NamespaceKind::Table && local == "table" {
        let name = attributes.get(&(NamespaceKind::Table, "name".to_owned())).cloned();
        if let Some(value) = &name { validate_text(value, "table:table table:name", false)?; }
        Some(ScopeFrame::Table(allocate_id(next_id)?, name))
    } else { None };
    Ok(scope)
}

#[allow(clippy::too_many_arguments)]
fn handle_empty(namespace: NamespaceKind, local: &str, attributes: Attributes, part: OdfNamedExpressionPart, depth: usize, parent: Option<&Frame>, group: &mut Option<ActiveGroup>, definition: &mut Option<ActiveDefinition>, seen: &mut HashSet<usize>, result: &mut OdfNamedExpressions) -> Result<()> {
    if definition.is_some() { return invalid("named range and expression elements must have empty content"); }
    if let Some(active) = group.as_mut() {
        if depth != active.parent_depth + 1 { return invalid("named definitions must be direct children of table:named-expressions"); }
        push_definition(active, parse_definition(namespace, local, attributes)?)?;
        return Ok(());
    }
    if namespace == NamespaceKind::Table && local == "named-expressions" {
        if !attributes.is_empty() { return invalid("table:named-expressions does not allow attributes"); }
        let (scope_id, scope) = named_scope(parent)?;
        if !seen.insert(scope_id) { return invalid("a spreadsheet or table may contain only one table:named-expressions group"); }
        result.groups.push(OdfNamedExpressionGroup { part, scope, definitions: Vec::new() });
        return Ok(());
    }
    if namespace == NamespaceKind::Table && matches!(local, "named-range" | "named-expression") { return invalid("named definitions must be direct children of table:named-expressions"); }
    Ok(())
}

fn parse_definition(namespace: NamespaceKind, local: &str, mut attributes: Attributes) -> Result<OdfNamedDefinition> {
    if namespace != NamespaceKind::Table { return invalid("table:named-expressions may contain only table:named-range or table:named-expression"); }
    let name = required(&mut attributes, "name")?;
    validate_text(&name, "table:name", false)?;
    let base_cell_address = attributes.remove(&(NamespaceKind::Table, "base-cell-address".to_owned())).map(OdfCellAddress::new).transpose()?;
    let value = match local {
        "named-range" => OdfNamedDefinition::Range {
            name,
            cell_range_address: OdfCellRangeAddress::new(required(&mut attributes, "cell-range-address")?)?,
            base_cell_address,
            usage: attributes.remove(&(NamespaceKind::Table, "range-usable-as".to_owned())).map(|value| OdfNamedRangeUsage::parse(&value)).transpose()?,
        },
        "named-expression" => OdfNamedDefinition::Expression {
            name,
            expression: OdfFormulaExpression::new(required(&mut attributes, "expression")?)?,
            base_cell_address,
        },
        _ => return invalid("table:named-expressions may contain only table:named-range or table:named-expression"),
    };
    if let Some(((namespace, local), _)) = attributes.into_iter().next() { return invalid(format!("unsupported {:?} named-definition attribute '{local}'", namespace)); }
    value.validate()?;
    Ok(value)
}

fn push_definition(group: &mut ActiveGroup, value: OdfNamedDefinition) -> Result<()> {
    if group.value.definitions.len() >= MAX_DEFINITIONS { return invalid(format!("named-expression group exceeds {MAX_DEFINITIONS} definitions")); }
    group.value.definitions.push(value);
    Ok(())
}

fn named_scope(parent: Option<&Frame>) -> Result<(usize, OdfNamedExpressionScope)> {
    match parent.and_then(|frame| frame.scope.as_ref()) {
        Some(ScopeFrame::Spreadsheet(id)) => Ok((*id, OdfNamedExpressionScope::Spreadsheet)),
        Some(ScopeFrame::Table(id, name)) => Ok((*id, OdfNamedExpressionScope::Table { name: name.clone() })),
        None => invalid("table:named-expressions must be a direct child of office:spreadsheet or table:table"),
    }
}

fn required(attributes: &mut Attributes, name: &str) -> Result<String> {
    attributes.remove(&(NamespaceKind::Table, name.to_owned())).ok_or_else(|| make_error(format!("named definition requires table:{name}")))
}

fn read_attributes<R: BufRead>(reader: &mut NsReader<R>, element: &BytesStart<'_>) -> Result<Attributes> {
    let mut result = HashMap::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| make_error(format!("invalid named-definition attribute: {error}")))?;
        let raw = attribute.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") { continue; }
        let (resolved, local) = reader.resolver_mut().resolve_attribute(attribute.key);
        let namespace = namespace_kind(&resolved)?;
        let local = decode(local.as_ref(), "attribute name")?;
        let value = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder()).map_err(|error| make_error(format!("invalid named-definition attribute value: {error}")))?.into_owned();
        validate_text(&value, &local, true)?;
        if result.insert((namespace, local.clone()), value).is_some() { return invalid(format!("duplicate named-definition attribute '{local}'")); }
    }
    Ok(result)
}

fn namespace_kind(resolved: &ResolveResult<'_>) -> Result<NamespaceKind> {
    match resolved {
        ResolveResult::Unbound => Ok(NamespaceKind::None),
        ResolveResult::Bound(namespace) => match namespace.as_ref() { OFFICE_NS => Ok(NamespaceKind::Office), TABLE_NS => Ok(NamespaceKind::Table), _ => Ok(NamespaceKind::Other) },
        ResolveResult::Unknown(prefix) => invalid(format!("unknown XML namespace prefix '{}'", String::from_utf8_lossy(prefix.as_ref()))),
    }
}

fn reject_spoofed_name(namespace: NamespaceKind, local: &str) -> Result<()> {
    if matches!(local, "named-expressions" | "named-range" | "named-expression") && namespace != NamespaceKind::Table { return invalid(format!("spoofed table:{local} element namespace")); }
    Ok(())
}

fn allocate_id(next: &mut usize) -> Result<usize> { let id = *next; *next = next.checked_add(1).ok_or_else(|| make_error("named-expression scope count overflow"))?; Ok(id) }
fn decode(value: &[u8], name: &str) -> Result<String> { std::str::from_utf8(value).map(str::to_owned).map_err(|error| make_error(format!("invalid UTF-8 {name}: {error}"))) }
fn validate_text(value: &str, name: &str, allow_empty: bool) -> Result<()> { if !allow_empty && value.is_empty() { return invalid(format!("{name} cannot be empty")); } if value.len() > MAX_VALUE_BYTES { return invalid(format!("{name} exceeds {MAX_VALUE_BYTES} bytes")); } if value.chars().any(|character| matches!(character, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}')) { return invalid(format!("{name} contains invalid XML characters")); } Ok(()) }
fn escape_attribute(output: &mut String, value: &str) { for character in value.chars() { match character { '&' => output.push_str("&amp;"), '<' => output.push_str("&lt;"), '"' => output.push_str("&quot;"), '\r' => output.push_str("&#13;"), '\n' => output.push_str("&#10;"), '\t' => output.push_str("&#9;"), _ => output.push(character) } } }
fn make_error(message: impl Into<String>) -> Error { Error::InvalidFormat(message.into()) }
fn invalid<T>(message: impl Into<String>) -> Result<T> { Err(make_error(message)) }

#[cfg(test)]
mod tests {
    use super::*;
    const PREFIX: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:spreadsheet>"#;
    const SUFFIX: &str = "</office:spreadsheet></office:body></office:document>";

    #[test]
    fn parses_scopes_usage_and_round_trips() {
        let xml = format!(r##"{PREFIX}<table:named-expressions><table:named-range table:name="Global" table:cell-range-address="$Sheet1.$A$1:.$B$2" table:base-cell-address="$Sheet1.$A$1" table:range-usable-as="print-range filter"/><table:named-expression table:name="Formula" table:expression="of:=SUM([.A1:.A2])"/></table:named-expressions><table:table table:name="Sheet1"><table:named-expressions><table:named-range table:name="Local" table:cell-range-address="#REF!" table:range-usable-as="none"></table:named-range></table:named-expressions></table:table>{SUFFIX}"##);
        let parsed = parse_named_expressions(&xml).unwrap();
        assert_eq!(parsed.groups.len(), 2);
        assert!(matches!(parsed.groups[0].definitions[0], OdfNamedDefinition::Range { usage: Some(OdfNamedRangeUsage::Uses(ref uses)), .. } if uses.len() == 2));
        assert_eq!(parsed.groups[1].scope, OdfNamedExpressionScope::Table { name: Some("Sheet1".into()) });
        let fragment = parsed.groups[0].to_xml_fragment().unwrap();
        let reparsed = parse_named_expressions(&format!("{PREFIX}{fragment}{SUFFIX}")).unwrap();
        assert_eq!(reparsed.groups[0].definitions, parsed.groups[0].definitions);
    }

    #[test]
    fn rejects_invalid_named_definition_grammar() {
        for body in [
            r#"<table:named-expressions><table:named-range table:name="x"/></table:named-expressions>"#,
            r#"<table:named-expressions><table:named-expression table:name="x" table:expression="1" bad="x"/></table:named-expressions>"#,
            r#"<table:named-expressions><table:named-range table:name="x" table:cell-range-address=".A1" table:range-usable-as="none filter"/></table:named-expressions>"#,
            r#"<table:named-expressions><table:named-range table:name="x" table:cell-range-address=".A1"/><table:named-expression table:name="x" table:expression="1"/></table:named-expressions>"#,
            r#"<table:named-expressions><table:named-range table:name="x" table:cell-range-address=".A1"><table:named-expression table:name="y" table:expression="1"/></table:named-range></table:named-expressions>"#,
        ] { assert!(parse_named_expressions(&format!("{PREFIX}{body}{SUFFIX}")).is_err(), "accepted {body}"); }
        assert!(parse_named_expressions(&format!("{PREFIX}<table:named-expressions/><table:named-expressions/>{SUFFIX}")).is_err());
        assert!(parse_named_expressions("<!DOCTYPE x><x/>").is_err());
    }

    #[test]
    fn parses_libreoffice_range_and_expression_fixtures_when_available() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../3rdparty/libreoffice-core/sc/qa/unit/data/functions");
        for (relative, expected) in [("financial/fods/yielddisc.fods", "days"), ("mathematical/fods/aggregate.fods", "columnOne")] {
            let Ok(xml) = std::fs::read_to_string(root.join(relative)) else { continue };
            let parsed = parse_named_expressions(&xml).unwrap();
            assert!(parsed.groups.iter().any(|group| group.get(expected).is_some()), "missing {expected} in {relative}");
        }
    }
}
