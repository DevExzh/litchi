//! Inert SpreadsheetML external data connection definitions.

use litchi_core::sheet::Result;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::collections::HashSet;

const X: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XS: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";
const CT: &str = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
const QUERY_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
const MAX: usize = 16 * 1024 * 1024;
const MAX_EXT: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_NODES: usize = 1_000_000;
const MAX_CONNECTIONS: usize = 65_536;
const MAX_PARAMETERS: usize = 65_536;
const MAX_TABLES: usize = 65_536;
const MAX_FIELDS: usize = 2001;
const MAX_STRING: usize = 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CredentialsMethod {
    Integrated,
    None,
    Stored,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HtmlFormatting {
    None,
    RichText,
    All,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextFileType {
    Mac,
    Windows,
    Dos,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextQualifier {
    DoubleQuote,
    SingleQuote,
    None,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TextFieldType {
    General,
    Text,
    MonthDayYear,
    DayMonthYear,
    YearMonthDay,
    MonthYearDay,
    DayYearMonth,
    YearDayMonth,
    Skip,
    EastAsianYearMonthDay,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ParameterType {
    Prompt,
    Value,
    Cell,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum WebTableSelector {
    Missing,
    String(String),
    Index(u32),
}
#[derive(Clone, PartialEq, Eq)]
pub struct DatabaseProperties {
    pub connection: String,
    pub command: Option<String>,
    pub server_command: Option<String>,
    pub command_type: Option<u32>,
}
#[derive(Clone, Debug, PartialEq, Eq, Default)]
pub struct OlapProperties {
    pub local: Option<bool>,
    pub local_connection: Option<String>,
    pub local_refresh: Option<bool>,
    pub send_locale: Option<bool>,
    pub row_drill_count: Option<u32>,
    pub server_fill: Option<bool>,
    pub server_number_format: Option<bool>,
    pub server_font: Option<bool>,
    pub server_font_color: Option<bool>,
}
#[derive(Clone, PartialEq, Eq, Default)]
pub struct WebQueryProperties {
    pub xml_source: Option<bool>,
    pub source_data: Option<bool>,
    pub parse_pre: Option<bool>,
    pub consecutive: Option<bool>,
    pub first_row: Option<bool>,
    pub excel97: Option<bool>,
    pub text_dates: Option<bool>,
    pub excel2000: Option<bool>,
    pub url: Option<String>,
    pub post: Option<String>,
    pub html_tables: Option<bool>,
    pub html_format: Option<HtmlFormatting>,
    pub edit_page: Option<String>,
    pub tables: Option<Vec<WebTableSelector>>,
}
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextField {
    pub field_type: Option<TextFieldType>,
    pub position: Option<u32>,
}
#[derive(Clone, PartialEq, Eq, Default)]
pub struct TextImportProperties {
    pub prompt: Option<bool>,
    pub file_type: Option<TextFileType>,
    pub code_page: Option<u32>,
    pub character_set: Option<String>,
    pub first_row: Option<u32>,
    pub source_file: Option<String>,
    pub delimited: Option<bool>,
    pub decimal: Option<String>,
    pub thousands: Option<String>,
    pub tab: Option<bool>,
    pub space: Option<bool>,
    pub comma: Option<bool>,
    pub semicolon: Option<bool>,
    pub consecutive: Option<bool>,
    pub qualifier: Option<TextQualifier>,
    pub delimiter: Option<String>,
    pub fields: Option<Vec<TextField>>,
}
#[derive(Clone, PartialEq)]
pub struct ConnectionParameter {
    pub name: Option<String>,
    pub sql_type: Option<i32>,
    pub parameter_type: Option<ParameterType>,
    pub refresh_on_change: Option<bool>,
    pub prompt: Option<String>,
    pub boolean: Option<bool>,
    pub double: Option<f64>,
    pub integer: Option<i32>,
    pub string: Option<String>,
    pub cell: Option<String>,
}
#[derive(Clone, PartialEq)]
pub struct Connection {
    pub id: u32,
    pub source_file: Option<String>,
    pub odc_file: Option<String>,
    pub keep_alive: Option<bool>,
    pub interval: Option<u32>,
    pub name: Option<String>,
    pub description: Option<String>,
    pub connection_type: Option<u32>,
    pub reconnection_method: Option<u32>,
    pub refreshed_version: u8,
    pub min_refreshable_version: Option<u8>,
    pub save_password: Option<bool>,
    pub new_connection: Option<bool>,
    pub deleted: Option<bool>,
    pub only_use_connection_file: Option<bool>,
    pub background: Option<bool>,
    pub refresh_on_load: Option<bool>,
    pub save_data: Option<bool>,
    pub credentials: Option<CredentialsMethod>,
    pub single_sign_on_id: Option<String>,
    pub database: Option<DatabaseProperties>,
    pub olap: Option<OlapProperties>,
    pub web: Option<WebQueryProperties>,
    pub text: Option<TextImportProperties>,
    pub parameters: Option<Vec<ConnectionParameter>>,
    pub extension_xml: Option<Vec<u8>>,
}
#[derive(Clone, Debug, PartialEq)]
pub struct Connections {
    pub connections: Vec<Connection>,
}

impl std::fmt::Debug for DatabaseProperties {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("DatabaseProperties")
            .field("connection", &"[REDACTED]")
            .field("command", &self.command.as_ref().map(|_| "[REDACTED]"))
            .field(
                "server_command",
                &self.server_command.as_ref().map(|_| "[REDACTED]"),
            )
            .field("command_type", &self.command_type)
            .finish()
    }
}

impl std::fmt::Debug for WebQueryProperties {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WebQueryProperties")
            .field("url", &self.url.as_ref().map(|_| "[REDACTED]"))
            .field("post", &self.post.as_ref().map(|_| "[REDACTED]"))
            .field("edit_page", &self.edit_page.as_ref().map(|_| "[REDACTED]"))
            .field("xml_source", &self.xml_source)
            .field("source_data", &self.source_data)
            .field("html_tables", &self.html_tables)
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for TextImportProperties {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("TextImportProperties")
            .field(
                "source_file",
                &self.source_file.as_ref().map(|_| "[REDACTED]"),
            )
            .field("file_type", &self.file_type)
            .field("code_page", &self.code_page)
            .field("first_row", &self.first_row)
            .field("delimited", &self.delimited)
            .finish_non_exhaustive()
    }
}

impl std::fmt::Debug for ConnectionParameter {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ConnectionParameter")
            .field("name", &self.name)
            .field("sql_type", &self.sql_type)
            .field("parameter_type", &self.parameter_type)
            .field("refresh_on_change", &self.refresh_on_change)
            .field("value", &"[REDACTED]")
            .finish()
    }
}

impl std::fmt::Debug for Connection {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Connection")
            .field("id", &self.id)
            .field("name", &self.name)
            .field("connection_type", &self.connection_type)
            .field("credentials", &self.credentials)
            .field(
                "source_file",
                &self.source_file.as_ref().map(|_| "[REDACTED]"),
            )
            .field("odc_file", &self.odc_file.as_ref().map(|_| "[REDACTED]"))
            .field(
                "single_sign_on_id",
                &self.single_sign_on_id.as_ref().map(|_| "[REDACTED]"),
            )
            .field("database", &self.database)
            .field("web", &self.web)
            .field("text", &self.text)
            .field("parameters", &self.parameters)
            .finish_non_exhaustive()
    }
}

impl Connections {
    pub fn parse(xml: &[u8]) -> Result<Self> {
        if xml.len() > MAX {
            return Err(invalid("connections part exceeds 16 MiB"));
        }
        let x = litchi_ooxml_common::mce::process_ooxml(xml)?;
        if x.len() > MAX {
            return Err(invalid("processed connections part exceeds 16 MiB"));
        }
        project(&parse_dom(x.as_ref())?)
    }
    pub fn to_xml(&self, strict: bool) -> Result<Vec<u8>> {
        validate(self)?;
        let ns = if strict { XS } else { X };
        let mut x = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><connections xmlns="{ns}">"#
        );
        for c in &self.connections {
            write_connection(&mut x, c, strict)?;
        }
        x.push_str("</connections>");
        if x.len() > MAX {
            return Err(invalid("serialized connections part exceeds 16 MiB"));
        }
        Ok(x.into_bytes())
    }
}

impl Connections {
    pub fn find(&self, id: u32) -> Option<&Connection> {
        self.connections
            .iter()
            .find(|connection| connection.id == id)
    }

    pub fn add(&mut self, connection: Connection) -> Result<()> {
        if self.find(connection.id).is_some() {
            return Err(invalid(format!(
                "duplicate connection ID {}",
                connection.id
            )));
        }
        self.connections.push(connection);
        validate(self)
    }

    pub fn update(&mut self, id: u32, connection: Connection) -> Result<()> {
        self.replace(id, connection)
    }

    pub fn replace(&mut self, id: u32, connection: Connection) -> Result<()> {
        if connection.id != id {
            return Err(invalid("replacement connection ID must remain stable"));
        }
        let offset = self
            .connections
            .iter()
            .position(|candidate| candidate.id == id)
            .ok_or_else(|| invalid(format!("connection ID {id} was not found")))?;
        let mut staged = self.clone();
        staged.connections[offset] = connection;
        validate(&staged)?;
        *self = staged;
        Ok(())
    }

    pub fn remove(&mut self, id: u32) -> Result<bool> {
        let Some(offset) = self
            .connections
            .iter()
            .position(|connection| connection.id == id)
        else {
            return Ok(false);
        };
        self.connections.remove(offset);
        Ok(true)
    }

    pub fn reorder(&mut self, ordered_ids: &[u32]) -> Result<()> {
        if ordered_ids.len() != self.connections.len() {
            return Err(invalid("connection reorder must contain every ID"));
        }
        let expected = self
            .connections
            .iter()
            .map(|item| item.id)
            .collect::<HashSet<_>>();
        let actual = ordered_ids.iter().copied().collect::<HashSet<_>>();
        if expected != actual || actual.len() != ordered_ids.len() {
            return Err(invalid("connection reorder is not a permutation"));
        }
        self.connections = ordered_ids
            .iter()
            .map(|id| self.find(*id).expect("permutation was validated").clone())
            .collect();
        Ok(())
    }
}

/// Store the complete inert connection set and validate every query-table reference first.
pub fn store_in_package(package: &mut OpcPackage, value: &Connections, strict: bool) -> Result<()> {
    store_in_package_with_query_table_validator(package, value, strict, query_table_connection_id)
}

/// Store connections while allowing the migration host to retain its complete
/// query-table parser for cross-part validation.
#[doc(hidden)]
pub fn store_in_package_with_query_table_validator<F>(
    package: &mut OpcPackage,
    value: &Connections,
    strict: bool,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let xml = value.to_xml(strict)?;
    validate_query_table_connection_ids(package, value, query_table_connection_id)?;
    let workbook_name = package.main_document_part()?.partname().clone();
    let existing = {
        let workbook = package.get_part(&workbook_name)?;
        let mut found = workbook
            .rels()
            .iter()
            .filter(|relationship| matches!(relationship.reltype(), REL | STRICT_REL));
        let first = found
            .next()
            .map(|relationship| {
                if relationship.is_external() {
                    return Err(invalid("connections relationship cannot be external"));
                }
                Ok((
                    relationship.r_id().to_string(),
                    relationship.target_partname()?,
                ))
            })
            .transpose()?;
        if found.next().is_some() {
            return Err(invalid("workbook has multiple connections relationships"));
        }
        first
    };
    if let Some((_, part_name)) = existing {
        let part = package.get_part(&part_name)?;
        if part.content_type() != CT {
            return Err(invalid(
                "existing connections part has invalid content type",
            ));
        }
        package.get_part_mut(&part_name)?.set_blob(xml);
    } else {
        let part_name = next_connections_part_name(package)?;
        let relationship_id = next_connections_relationship_id(package, &workbook_name)?;
        package.try_add_part(Box::new(litchi_opc::part::BlobPart::new(
            part_name.clone(),
            CT.into(),
            xml,
        )))?;
        package
            .get_part_mut(&workbook_name)?
            .rels_mut()
            .add_relationship(
                if strict { STRICT_REL } else { REL }.into(),
                part_name.relative_ref(workbook_name.base_uri()),
                relationship_id,
                false,
            );
    }
    package.unsign();
    Ok(())
}

pub fn remove_from_package(package: &mut OpcPackage) -> Result<bool> {
    if package
        .iter_parts()
        .any(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        return Err(invalid(
            "cannot remove connections while query-table parts remain",
        ));
    }
    let workbook_name = package.main_document_part()?.partname().clone();
    let relationship = package
        .get_part(&workbook_name)?
        .rels()
        .iter()
        .find(|relationship| matches!(relationship.reltype(), REL | STRICT_REL))
        .map(|relationship| {
            relationship
                .target_partname()
                .map(|part_name| (relationship.r_id().to_string(), part_name))
        })
        .transpose()?;
    let Some((relationship_id, part_name)) = relationship else {
        return Ok(false);
    };
    package
        .get_part_mut(&workbook_name)?
        .rels_mut()
        .remove(&relationship_id);
    if !package_part_is_referenced(package, &part_name) {
        package.remove_part(&part_name);
    }
    package.unsign();
    Ok(true)
}

fn validate_query_table_connection_ids<F>(
    package: &OpcPackage,
    value: &Connections,
    query_table_connection_id: F,
) -> Result<()>
where
    F: Fn(&[u8]) -> Result<u32>,
{
    let ids = value
        .connections
        .iter()
        .map(|connection| connection.id)
        .collect::<HashSet<_>>();
    for part in package
        .iter_parts()
        .filter(|part| part.content_type() == QUERY_TABLE_CONTENT_TYPE)
    {
        let connection_id = query_table_connection_id(part.blob())?;
        if !ids.contains(&connection_id) {
            return Err(invalid(format!(
                "query-table part '{}' references missing connection ID {}",
                part.partname(),
                connection_id
            )));
        }
    }
    Ok(())
}

fn query_table_connection_id(xml: &[u8]) -> Result<u32> {
    if xml.len() > 8 * 1024 * 1024 {
        return Err(invalid("query-table part exceeds 8 MiB"));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    if processed.len() > 8 * 1024 * 1024 {
        return Err(invalid("processed query-table part exceeds 8 MiB"));
    }
    let root = parse_dom(processed.as_ref())?;
    expect(&root, "queryTable")?;
    let _name = req(&root, "name")?;
    let connection_id = u32req(&root, "connectionId")?;
    only_unqualified(
        &root,
        &[
            "name",
            "headers",
            "rowNumbers",
            "disableRefresh",
            "backgroundRefresh",
            "firstBackgroundRefresh",
            "refreshOnLoad",
            "growShrinkType",
            "fillFormulas",
            "removeDataOnSave",
            "disableEdit",
            "preserveFormatting",
            "adjustColumnWidth",
            "intermediate",
            "connectionId",
            "autoFormatId",
            "applyNumberFormats",
            "applyBorderFormats",
            "applyFontFormats",
            "applyPatternFormats",
            "applyAlignmentFormats",
            "applyWidthHeightFormats",
        ],
    )?;
    kids(&root)?;
    Ok(connection_id)
}

fn next_connections_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 0..=65_536u32 {
        let name = if suffix == 0 {
            "/xl/connections.xml".into()
        } else {
            format!("/xl/connections{suffix}.xml")
        };
        let candidate = PackURI::new(&name)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections part name"))
}

fn next_connections_relationship_id(package: &OpcPackage, workbook: &PackURI) -> Result<String> {
    let relationships = package.get_part(workbook)?.rels();
    for suffix in 1..=65_537u32 {
        let candidate = format!("rIdConnections{suffix}");
        if relationships.get(&candidate).is_none() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free connections relationship ID"))
}

fn package_part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}
pub fn load_from_package(package: &OpcPackage) -> Result<Option<Connections>> {
    let workbook = package.main_document_part()?;
    let mut found = workbook
        .rels()
        .iter()
        .filter(|x| matches!(x.reltype(), REL | STRICT_REL));
    let Some(rel) = found.next() else {
        return Ok(None);
    };
    if found.next().is_some() {
        return Err(invalid("workbook has multiple connections relationships"));
    }
    if rel.is_external() {
        return Err(invalid("connections relationship cannot be external"));
    }
    let uri: PackURI = rel.target_partname()?;
    let part = package.get_part(&uri)?;
    if part.content_type() != CT {
        return Err(invalid(format!(
            "connections part '{uri}' has invalid content type '{}'",
            part.content_type()
        )));
    }
    if part.rels().iter().next().is_some() {
        return Err(invalid("connections part must not have relationships"));
    }
    Ok(Some(Connections::parse(part.blob())?))
}

#[derive(Clone)]
struct Attr {
    q: String,
    ns: String,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
struct Node {
    q: String,
    ns: String,
    l: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
}
fn parse_dom(xml: &[u8]) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0;
    loop {
        let d = rd.decoder();
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count += 1;
                if count > MAX_NODES || stack.len() >= MAX_DEPTH {
                    return Err(invalid("connections XML resource limit exceeded"));
                }
                stack.push(make(&e, d, &stack)?);
            },
            Ok(Event::Empty(e)) => {
                count += 1;
                if count > MAX_NODES {
                    return Err(invalid("connections node limit exceeded"));
                }
                let n = make(&e, d, &stack)?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v))
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside connections"));
                }
            },
            Ok(Event::CData(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content
                        .push(Content::CData(t.decode().map_err(xml_error)?.into_owned()))
                } else {
                    return Err(invalid("CDATA outside connections"));
                }
            },
            Ok(Event::Comment(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(
                        t.decode().map_err(xml_error)?.into_owned(),
                    ))
                }
            },
            Ok(Event::GeneralRef(t)) => {
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(
                        litchi_ooxml_common::xml::decode_xml_reference(&t)?,
                    ))
                } else {
                    return Err(invalid("entity outside connections"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated connections XML"));
    }
    root.ok_or_else(|| invalid("missing connections root"))
}
fn make(e: &BytesStart<'_>, d: Decoder, stack: &[Node]) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|x| x.bindings.clone()).unwrap_or_default();
    let mut raw = Vec::new();
    for a in e.attributes().with_checks(true) {
        let a = a.map_err(xml_error)?;
        raw.push((
            std::str::from_utf8(a.key.as_ref())
                .map_err(xml_error)?
                .to_string(),
            a.decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
                .map_err(xml_error)?
                .into_owned(),
        ));
    }
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(x) = bindings.iter_mut().find(|x| x.0 == key) {
                x.1 = v.clone()
            } else {
                bindings.push((key, v.clone()))
            }
        }
    }
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    let ns = resolve(&bindings, pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            String::new()
        } else {
            resolve(&bindings, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content.push(Content::Node(n))
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn project(n: &Node) -> Result<Connections> {
    expect(n, "connections")?;
    noattrs(n)?;
    let mut out = Vec::new();
    for c in kids(n)? {
        if out.len() >= MAX_CONNECTIONS {
            return Err(invalid("connection limit exceeded"));
        }
        out.push(parse_connection(c)?);
    }
    if out.is_empty() {
        return Err(invalid("connections requires at least one connection"));
    }
    let value = Connections { connections: out };
    validate(&value)?;
    Ok(value)
}
fn parse_connection(n: &Node) -> Result<Connection> {
    expect(n, "connection")?;
    let mut c = Connection {
        id: u32req(n, "id")?,
        source_file: aopt(n, "sourceFile")?,
        odc_file: aopt(n, "odcFile")?,
        keep_alive: bopt(n, "keepAlive")?,
        interval: u32opt(n, "interval")?,
        name: aopt(n, "name")?,
        description: aopt(n, "description")?,
        connection_type: u32opt(n, "type")?,
        reconnection_method: u32opt(n, "reconnectionMethod")?,
        refreshed_version: u8req(n, "refreshedVersion")?,
        min_refreshable_version: u8opt(n, "minRefreshableVersion")?,
        save_password: bopt(n, "savePassword")?,
        new_connection: bopt(n, "new")?,
        deleted: bopt(n, "deleted")?,
        only_use_connection_file: bopt(n, "onlyUseConnectionFile")?,
        background: bopt(n, "background")?,
        refresh_on_load: bopt(n, "refreshOnLoad")?,
        save_data: bopt(n, "saveData")?,
        credentials: aopt(n, "credentials")?.map(parse_credentials).transpose()?,
        single_sign_on_id: aopt(n, "singleSignOnId")?,
        database: None,
        olap: None,
        web: None,
        text: None,
        parameters: None,
        extension_xml: None,
    };
    only(
        n,
        &[
            "id",
            "sourceFile",
            "odcFile",
            "keepAlive",
            "interval",
            "name",
            "description",
            "type",
            "reconnectionMethod",
            "refreshedVersion",
            "minRefreshableVersion",
            "savePassword",
            "new",
            "deleted",
            "onlyUseConnectionFile",
            "background",
            "refreshOnLoad",
            "saveData",
            "credentials",
            "singleSignOnId",
        ],
    )?;
    let mut order = 0;
    for child in kids(n)? {
        expect_any(child)?;
        let i = match child.l.as_str() {
            "dbPr" => 0,
            "olapPr" => 1,
            "webPr" => 2,
            "textPr" => 3,
            "parameters" => 4,
            "extLst" => 5,
            _ => return Err(invalid("unexpected connection child")),
        };
        if i < order {
            return Err(invalid("connection children out of order"));
        }
        order = i;
        match i {
            0 => set(&mut c.database, parse_db(child)?)?,
            1 => set(&mut c.olap, parse_olap(child)?)?,
            2 => set(&mut c.web, parse_web(child)?)?,
            3 => set(&mut c.text, parse_text(child)?)?,
            4 => set(&mut c.parameters, parse_parameters(child)?)?,
            5 => set(&mut c.extension_xml, node_xml(child, false)?)?,
            _ => return Err(invalid("unexpected connection child index")),
        }
    }
    Ok(c)
}
fn parse_db(n: &Node) -> Result<DatabaseProperties> {
    let v = DatabaseProperties {
        connection: req(n, "connection")?,
        command: aopt(n, "command")?,
        server_command: aopt(n, "serverCommand")?,
        command_type: u32opt(n, "commandType")?,
    };
    only(
        n,
        &["connection", "command", "serverCommand", "commandType"],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_olap(n: &Node) -> Result<OlapProperties> {
    let v = OlapProperties {
        local: bopt(n, "local")?,
        local_connection: aopt(n, "localConnection")?,
        local_refresh: bopt(n, "localRefresh")?,
        send_locale: bopt(n, "sendLocale")?,
        row_drill_count: u32opt(n, "rowDrillCount")?,
        server_fill: bopt(n, "serverFill")?,
        server_number_format: bopt(n, "serverNumberFormat")?,
        server_font: bopt(n, "serverFont")?,
        server_font_color: bopt(n, "serverFontColor")?,
    };
    only(
        n,
        &[
            "local",
            "localConnection",
            "localRefresh",
            "sendLocale",
            "rowDrillCount",
            "serverFill",
            "serverNumberFormat",
            "serverFont",
            "serverFontColor",
        ],
    )?;
    leaf(n)?;
    Ok(v)
}
fn parse_web(n: &Node) -> Result<WebQueryProperties> {
    let mut v = WebQueryProperties {
        xml_source: bopt(n, "xml")?,
        source_data: bopt(n, "sourceData")?,
        parse_pre: bopt(n, "parsePre")?,
        consecutive: bopt(n, "consecutive")?,
        first_row: bopt(n, "firstRow")?,
        excel97: bopt(n, "xl97")?,
        text_dates: bopt(n, "textDates")?,
        excel2000: bopt(n, "xl2000")?,
        url: aopt(n, "url")?,
        post: aopt(n, "post")?,
        html_tables: bopt(n, "htmlTables")?,
        html_format: aopt(n, "htmlFormat")?.map(parse_html).transpose()?,
        edit_page: aopt(n, "editPage")?,
        tables: None,
    };
    only(
        n,
        &[
            "xml",
            "sourceData",
            "parsePre",
            "consecutive",
            "firstRow",
            "xl97",
            "textDates",
            "xl2000",
            "url",
            "post",
            "htmlTables",
            "htmlFormat",
            "editPage",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("webPr permits one tables child"));
    }
    if let Some(t) = c.first() {
        expect(t, "tables")?;
        v.tables = Some(parse_tables(t)?);
    }
    Ok(v)
}
fn parse_tables(n: &Node) -> Result<Vec<WebTableSelector>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for c in kids(n)? {
        if out.len() >= MAX_TABLES {
            return Err(invalid("web table selector limit exceeded"));
        }
        expect_any(c)?;
        out.push(match c.l.as_str() {
            "m" => {
                noattrs(c)?;
                leaf(c)?;
                WebTableSelector::Missing
            },
            "s" => {
                let v = req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::String(v)
            },
            "x" => {
                let v = u32req(c, "v")?;
                only(c, &["v"])?;
                leaf(c)?;
                WebTableSelector::Index(v)
            },
            _ => return Err(invalid("invalid web table selector")),
        });
    }
    if out.is_empty() {
        return Err(invalid("tables requires a selector"));
    }
    check_count(count, out.len(), "tables")?;
    Ok(out)
}
fn parse_text(n: &Node) -> Result<TextImportProperties> {
    let mut v = TextImportProperties {
        prompt: bopt(n, "prompt")?,
        file_type: aopt(n, "fileType")?.map(parse_file).transpose()?,
        code_page: u32opt(n, "codePage")?,
        character_set: aopt(n, "characterSet")?,
        first_row: u32opt(n, "firstRow")?,
        source_file: aopt(n, "sourceFile")?,
        delimited: bopt(n, "delimited")?,
        decimal: aopt(n, "decimal")?,
        thousands: aopt(n, "thousands")?,
        tab: bopt(n, "tab")?,
        space: bopt(n, "space")?,
        comma: bopt(n, "comma")?,
        semicolon: bopt(n, "semicolon")?,
        consecutive: bopt(n, "consecutive")?,
        qualifier: aopt(n, "qualifier")?.map(parse_qualifier).transpose()?,
        delimiter: aopt(n, "delimiter")?,
        fields: None,
    };
    only(
        n,
        &[
            "prompt",
            "fileType",
            "codePage",
            "characterSet",
            "firstRow",
            "sourceFile",
            "delimited",
            "decimal",
            "thousands",
            "tab",
            "space",
            "comma",
            "semicolon",
            "consecutive",
            "qualifier",
            "delimiter",
        ],
    )?;
    let c = kids(n)?;
    if c.len() > 1 {
        return Err(invalid("textPr permits one textFields child"));
    }
    if let Some(f) = c.first() {
        expect(f, "textFields")?;
        let count = u32opt(f, "count")?;
        only(f, &["count"])?;
        let mut fields = Vec::new();
        for e in kids(f)? {
            if fields.len() >= MAX_FIELDS {
                return Err(invalid("text field limit exceeded"));
            }
            expect(e, "textField")?;
            fields.push(TextField {
                field_type: aopt(e, "type")?.map(parse_field).transpose()?,
                position: u32opt(e, "position")?,
            });
            only(e, &["type", "position"])?;
            leaf(e)?;
        }
        if fields.is_empty() {
            return Err(invalid("textFields requires a textField"));
        }
        check_count(count, fields.len(), "textFields")?;
        v.fields = Some(fields);
    }
    Ok(v)
}
fn parse_parameters(n: &Node) -> Result<Vec<ConnectionParameter>> {
    let count = u32opt(n, "count")?;
    only(n, &["count"])?;
    let mut out = Vec::new();
    for p in kids(n)? {
        if out.len() >= MAX_PARAMETERS {
            return Err(invalid("parameter limit exceeded"));
        }
        expect(p, "parameter")?;
        let double = match aopt(p, "double")? {
            Some(x) => {
                let v = x
                    .parse::<f64>()
                    .map_err(|_| invalid("invalid parameter double"))?;
                if !v.is_finite() {
                    return Err(invalid("non-finite parameter double"));
                }
                Some(v)
            },
            None => None,
        };
        out.push(ConnectionParameter {
            name: aopt(p, "name")?,
            sql_type: i32opt(p, "sqlType")?,
            parameter_type: aopt(p, "parameterType")?
                .map(parse_parameter_type)
                .transpose()?,
            refresh_on_change: bopt(p, "refreshOnChange")?,
            prompt: aopt(p, "prompt")?,
            boolean: bopt(p, "boolean")?,
            double,
            integer: i32opt(p, "integer")?,
            string: aopt(p, "string")?,
            cell: aopt(p, "cell")?,
        });
        only(
            p,
            &[
                "name",
                "sqlType",
                "parameterType",
                "refreshOnChange",
                "prompt",
                "boolean",
                "double",
                "integer",
                "string",
                "cell",
            ],
        )?;
        leaf(p)?;
    }
    if out.is_empty() {
        return Err(invalid("parameters requires a parameter"));
    }
    check_count(count, out.len(), "parameters")?;
    Ok(out)
}

fn write_connection(x: &mut String, c: &Connection, s: bool) -> Result<()> {
    x.push_str("<connection");
    num(x, "id", c.id);
    str_opt(x, "sourceFile", c.source_file.as_deref());
    str_opt(x, "odcFile", c.odc_file.as_deref());
    bool_opt(x, "keepAlive", c.keep_alive);
    num_opt(x, "interval", c.interval);
    str_opt(x, "name", c.name.as_deref());
    str_opt(x, "description", c.description.as_deref());
    num_opt(x, "type", c.connection_type);
    num_opt(x, "reconnectionMethod", c.reconnection_method);
    num(x, "refreshedVersion", c.refreshed_version);
    num_opt(x, "minRefreshableVersion", c.min_refreshable_version);
    bool_opt(x, "savePassword", c.save_password);
    bool_opt(x, "new", c.new_connection);
    bool_opt(x, "deleted", c.deleted);
    bool_opt(x, "onlyUseConnectionFile", c.only_use_connection_file);
    bool_opt(x, "background", c.background);
    bool_opt(x, "refreshOnLoad", c.refresh_on_load);
    bool_opt(x, "saveData", c.save_data);
    if let Some(v) = c.credentials {
        attr(x, "credentials", credentials_str(v));
    }
    str_opt(x, "singleSignOnId", c.single_sign_on_id.as_deref());
    if c.database.is_none()
        && c.olap.is_none()
        && c.web.is_none()
        && c.text.is_none()
        && c.parameters.is_none()
        && c.extension_xml.is_none()
    {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    if let Some(v) = &c.database {
        x.push_str("<dbPr");
        attr(x, "connection", &v.connection);
        str_opt(x, "command", v.command.as_deref());
        str_opt(x, "serverCommand", v.server_command.as_deref());
        num_opt(x, "commandType", v.command_type);
        x.push_str("/>");
    }
    if let Some(v) = &c.olap {
        write_olap(x, v);
    }
    if let Some(v) = &c.web {
        write_web(x, v);
    }
    if let Some(v) = &c.text {
        write_text(x, v);
    }
    if let Some(v) = &c.parameters {
        write_parameters(x, v);
    }
    if let Some(v) = &c.extension_xml {
        opaque(x, v, s)?;
    }
    x.push_str("</connection>");
    Ok(())
}
fn write_olap(x: &mut String, v: &OlapProperties) {
    x.push_str("<olapPr");
    for (n, b) in [
        ("local", v.local),
        ("localRefresh", v.local_refresh),
        ("sendLocale", v.send_locale),
        ("serverFill", v.server_fill),
        ("serverNumberFormat", v.server_number_format),
        ("serverFont", v.server_font),
        ("serverFontColor", v.server_font_color),
    ] {
        bool_opt(x, n, b)
    }
    str_opt(x, "localConnection", v.local_connection.as_deref());
    num_opt(x, "rowDrillCount", v.row_drill_count);
    x.push_str("/>")
}
fn write_web(x: &mut String, v: &WebQueryProperties) {
    x.push_str("<webPr");
    for (n, b) in [
        ("xml", v.xml_source),
        ("sourceData", v.source_data),
        ("parsePre", v.parse_pre),
        ("consecutive", v.consecutive),
        ("firstRow", v.first_row),
        ("xl97", v.excel97),
        ("textDates", v.text_dates),
        ("xl2000", v.excel2000),
        ("htmlTables", v.html_tables),
    ] {
        bool_opt(x, n, b)
    }
    str_opt(x, "url", v.url.as_deref());
    str_opt(x, "post", v.post.as_deref());
    if let Some(h) = v.html_format {
        attr(x, "htmlFormat", html_str(h));
    }
    str_opt(x, "editPage", v.edit_page.as_deref());
    if let Some(t) = &v.tables {
        x.push_str("><tables");
        num(x, "count", t.len());
        x.push('>');
        for z in t {
            match z {
                WebTableSelector::Missing => x.push_str("<m/>"),
                WebTableSelector::String(v) => {
                    x.push_str("<s");
                    attr(x, "v", v);
                    x.push_str("/>");
                },
                WebTableSelector::Index(v) => {
                    x.push_str("<x");
                    num(x, "v", *v);
                    x.push_str("/>");
                },
            }
        }
        x.push_str("</tables></webPr>");
    } else {
        x.push_str("/>");
    }
}
fn write_text(x: &mut String, v: &TextImportProperties) {
    x.push_str("<textPr");
    bool_opt(x, "prompt", v.prompt);
    if let Some(z) = v.file_type {
        attr(x, "fileType", file_str(z));
    }
    num_opt(x, "codePage", v.code_page);
    str_opt(x, "characterSet", v.character_set.as_deref());
    num_opt(x, "firstRow", v.first_row);
    str_opt(x, "sourceFile", v.source_file.as_deref());
    for (n, b) in [
        ("delimited", v.delimited),
        ("tab", v.tab),
        ("space", v.space),
        ("comma", v.comma),
        ("semicolon", v.semicolon),
        ("consecutive", v.consecutive),
    ] {
        bool_opt(x, n, b)
    }
    str_opt(x, "decimal", v.decimal.as_deref());
    str_opt(x, "thousands", v.thousands.as_deref());
    if let Some(z) = v.qualifier {
        attr(x, "qualifier", qualifier_str(z));
    }
    str_opt(x, "delimiter", v.delimiter.as_deref());
    if let Some(f) = &v.fields {
        x.push_str("><textFields");
        num(x, "count", f.len());
        x.push('>');
        for z in f {
            x.push_str("<textField");
            if let Some(t) = z.field_type {
                attr(x, "type", field_str(t));
            }
            num_opt(x, "position", z.position);
            x.push_str("/>");
        }
        x.push_str("</textFields></textPr>");
    } else {
        x.push_str("/>");
    }
}
fn write_parameters(x: &mut String, v: &[ConnectionParameter]) {
    x.push_str("<parameters");
    num(x, "count", v.len());
    x.push('>');
    for p in v {
        x.push_str("<parameter");
        str_opt(x, "name", p.name.as_deref());
        num_opt(x, "sqlType", p.sql_type);
        if let Some(z) = p.parameter_type {
            attr(x, "parameterType", parameter_str(z));
        }
        bool_opt(x, "refreshOnChange", p.refresh_on_change);
        str_opt(x, "prompt", p.prompt.as_deref());
        bool_opt(x, "boolean", p.boolean);
        if let Some(z) = p.double {
            attr(x, "double", &z.to_string());
        }
        num_opt(x, "integer", p.integer);
        str_opt(x, "string", p.string.as_deref());
        str_opt(x, "cell", p.cell.as_deref());
        x.push_str("/>");
    }
    x.push_str("</parameters>");
}

fn validate(v: &Connections) -> Result<()> {
    if v.connections.is_empty() || v.connections.len() > MAX_CONNECTIONS {
        return Err(invalid("invalid connection count"));
    }
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    let mut ext = 0usize;
    for c in &v.connections {
        if !ids.insert(c.id) {
            return Err(invalid("duplicate connection id"));
        }
        if let Some(n) = &c.name {
            bounded(n)?;
            if !names.insert(n) {
                return Err(invalid("duplicate connection name"));
            }
        }
        for z in [
            c.source_file.as_deref(),
            c.odc_file.as_deref(),
            c.description.as_deref(),
            c.single_sign_on_id.as_deref(),
            c.database.as_ref().map(|x| x.connection.as_str()),
            c.database.as_ref().and_then(|x| x.command.as_deref()),
            c.database
                .as_ref()
                .and_then(|x| x.server_command.as_deref()),
            c.olap.as_ref().and_then(|x| x.local_connection.as_deref()),
            c.web.as_ref().and_then(|x| x.url.as_deref()),
            c.web.as_ref().and_then(|x| x.post.as_deref()),
            c.web.as_ref().and_then(|x| x.edit_page.as_deref()),
            c.text.as_ref().and_then(|x| x.source_file.as_deref()),
        ]
        .into_iter()
        .flatten()
        {
            bounded(z)?;
        }
        if c.web
            .as_ref()
            .and_then(|x| x.tables.as_ref())
            .is_some_and(|x| x.is_empty() || x.len() > MAX_TABLES)
        {
            return Err(invalid("invalid web table count"));
        }
        if c.text
            .as_ref()
            .and_then(|x| x.fields.as_ref())
            .is_some_and(|x| x.is_empty() || x.len() > MAX_FIELDS)
        {
            return Err(invalid("invalid text field count"));
        }
        if c.parameters
            .as_ref()
            .is_some_and(|x| x.is_empty() || x.len() > MAX_PARAMETERS)
        {
            return Err(invalid("invalid parameter count"));
        }
        if let Some(e) = &c.extension_xml {
            parse_dom(e)?;
            ext = ext
                .checked_add(e.len())
                .ok_or_else(|| invalid("extension size overflow"))?;
        }
    }
    if ext > MAX_EXT {
        return Err(invalid("connection extensions exceed 8 MiB"));
    }
    Ok(())
}
fn opaque(x: &mut String, b: &[u8], strict: bool) -> Result<()> {
    parse_dom(b)?;
    let mut s = std::str::from_utf8(b).map_err(xml_error)?.to_string();
    if strict {
        s = s.replace(X, XS)
    } else {
        s = s.replace(XS, X)
    }
    x.push_str(&s);
    Ok(())
}
fn node_xml(n: &Node, s: bool) -> Result<Vec<u8>> {
    let mut x = String::new();
    node_write(&mut x, n, s)?;
    Ok(x.into_bytes())
}
fn node_write(x: &mut String, n: &Node, s: bool) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in &n.bindings {
        if p.is_empty() {
            x.push_str(" xmlns=\"")
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"")
        }
        esc(
            x,
            if s && u == X {
                XS
            } else if !s && u == XS {
                X
            } else {
                u
            },
        );
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n, s)?,
            Content::Text(v) => text_escape(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}

fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed connections")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node) -> Result<()> {
    if kids(n)?.is_empty() {
        Ok(())
    } else {
        Err(invalid("connection leaf has children"))
    }
}
fn expect(n: &Node, l: &str) -> Result<()> {
    if (n.ns == X || n.ns == XS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected SpreadsheetML {l}")))
    }
}
fn expect_any(n: &Node) -> Result<()> {
    if n.ns == X || n.ns == XS {
        Ok(())
    } else {
        Err(invalid("expected SpreadsheetML child"))
    }
}
fn aopt(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns.is_empty() && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            bounded(&a.v)?;
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
fn req(n: &Node, l: &str) -> Result<String> {
    aopt(n, l)?.ok_or_else(|| invalid(format!("missing required attribute '{l}'")))
}
fn bopt(n: &Node, l: &str) -> Result<Option<bool>> {
    match aopt(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn u32opt(n: &Node, l: &str) -> Result<Option<u32>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid u32 '{l}'"))))
        .transpose()
}
fn u32req(n: &Node, l: &str) -> Result<u32> {
    u32opt(n, l)?.ok_or_else(|| invalid(format!("missing u32 '{l}'")))
}
fn u8opt(n: &Node, l: &str) -> Result<Option<u8>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid u8 '{l}'"))))
        .transpose()
}
fn u8req(n: &Node, l: &str) -> Result<u8> {
    u8opt(n, l)?.ok_or_else(|| invalid(format!("missing u8 '{l}'")))
}
fn i32opt(n: &Node, l: &str) -> Result<Option<i32>> {
    aopt(n, l)?
        .map(|x| x.parse().map_err(|_| invalid(format!("invalid i32 '{l}'"))))
        .transpose()
}
fn only(n: &Node, a: &[&str]) -> Result<()> {
    for x in &n.attrs {
        if !x.ns.is_empty() || !a.contains(&x.l.as_str()) {
            return Err(invalid(format!("unexpected attribute '{}'", x.q)));
        }
    }
    Ok(())
}

fn only_unqualified(n: &Node, allowed: &[&str]) -> Result<()> {
    for attribute in &n.attrs {
        if attribute.ns.is_empty() && !allowed.contains(&attribute.l.as_str()) {
            return Err(invalid(format!("unexpected attribute '{}'", attribute.q)));
        }
    }
    Ok(())
}
fn noattrs(n: &Node) -> Result<()> {
    only(n, &[])
}
fn set<T>(s: &mut Option<T>, v: T) -> Result<()> {
    if s.replace(v).is_some() {
        Err(invalid("duplicate connection property"))
    } else {
        Ok(())
    }
}
fn check_count(c: Option<u32>, actual: usize, n: &str) -> Result<()> {
    if c.is_some_and(|x| x as usize != actual) {
        Err(invalid(format!("{n} count mismatch")))
    } else {
        Ok(())
    }
}
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(b: &[(String, String)], p: &str) -> Result<String> {
    if p == "xml" {
        return Ok("http://www.w3.org/XML/1998/namespace".into());
    }
    b.iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
fn bounded(v: &str) -> Result<()> {
    if v.len() > MAX_STRING {
        Err(invalid("connection string exceeds 1 MiB"))
    } else {
        Ok(())
    }
}
fn attr(x: &mut String, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    esc(x, v);
    x.push('"')
}
fn str_opt(x: &mut String, n: &str, v: Option<&str>) {
    if let Some(v) = v {
        attr(x, n, v)
    }
}
fn bool_opt(x: &mut String, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr(x, n, if v { "1" } else { "0" })
    }
}
fn num<T: std::fmt::Display>(x: &mut String, n: &str, v: T) {
    attr(x, n, &v.to_string())
}
fn num_opt<T: std::fmt::Display>(x: &mut String, n: &str, v: Option<T>) {
    if let Some(v) = v {
        num(x, n, v)
    }
}
fn esc(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn text_escape(x: &mut String, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! en{($p:ident,$w:ident,$t:ty,$($s:literal=>$v:path),+)=>{fn $p(s:String)->Result<$t>{match s.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{s}'")))}}fn $w(v:$t)->&'static str{match v{$($v=>$s,)+}}}}
en!(parse_credentials,credentials_str,CredentialsMethod,"integrated"=>CredentialsMethod::Integrated,"none"=>CredentialsMethod::None,"stored"=>CredentialsMethod::Stored);
en!(parse_html,html_str,HtmlFormatting,"none"=>HtmlFormatting::None,"rtf"=>HtmlFormatting::RichText,"all"=>HtmlFormatting::All);
en!(parse_file,file_str,TextFileType,"mac"=>TextFileType::Mac,"win"=>TextFileType::Windows,"dos"=>TextFileType::Dos);
en!(parse_qualifier,qualifier_str,TextQualifier,"doubleQuote"=>TextQualifier::DoubleQuote,"singleQuote"=>TextQualifier::SingleQuote,"none"=>TextQualifier::None);
en!(parse_parameter_type,parameter_str,ParameterType,"prompt"=>ParameterType::Prompt,"value"=>ParameterType::Value,"cell"=>ParameterType::Cell);
en!(parse_field,field_str,TextFieldType,"general"=>TextFieldType::General,"text"=>TextFieldType::Text,"MDY"=>TextFieldType::MonthDayYear,"DMY"=>TextFieldType::DayMonthYear,"YMD"=>TextFieldType::YearMonthDay,"MYD"=>TextFieldType::MonthYearDay,"DYM"=>TextFieldType::DayYearMonth,"YDM"=>TextFieldType::YearDayMonth,"skip"=>TextFieldType::Skip,"EMD"=>TextFieldType::EastAsianYearMonthDay);
fn invalid(v: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, v.into()).into()
}
fn xml_error(e: impl std::fmt::Display) -> Box<dyn std::error::Error + Send + Sync> {
    invalid(e.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::phys_pkg::{PhysPkgReader, PhysPkgWriter};
    use litchi_opc::{BlobPart, PackURI, Part};
    fn f(b: &[u8]) -> Connections {
        let p = OpcPackage::from_bytes(b).unwrap();
        load_from_package(&p).unwrap().unwrap()
    }
    fn f_without_broken_thumbnail(b: &[u8]) -> Connections {
        let reader = PhysPkgReader::new(b).unwrap();
        let mut writer = PhysPkgWriter::new();
        for name in reader.member_names().unwrap() {
            if name == "docProps/thumbnail.jpeg" {
                continue;
            }
            let uri = PackURI::new(format!("/{name}")).unwrap();
            let mut data = reader.blob_for(&uri).unwrap();
            if name == "_rels/.rels" {
                let xml = String::from_utf8(data).unwrap();
                data = xml.replace("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\" Target=\"docProps/thumbnail.jpeg\"/>", "").into_bytes();
            }
            writer.write(&uri, &data).unwrap();
        }
        f(&writer.finish().unwrap())
    }
    #[test]
    fn poi_web_paths_are_inert() {
        let v = f(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/56169.xlsx"
        ));
        assert_eq!(v.connections.len(), 3);
        assert!(
            v.connections[0]
                .web
                .as_ref()
                .unwrap()
                .url
                .as_ref()
                .unwrap()
                .starts_with("\\\\snb.ch")
        );
    }
    #[test]
    fn poi_database_mce_and_strict_roundtrip() {
        let v = f(include_bytes!(
            "../../../test-data/poi/test-data/spreadsheet/ExcelPivotTableSample.xlsx"
        ));
        let db = v.connections[0].database.as_ref().unwrap();
        assert!(db.connection.contains("Microsoft.ACE.OLEDB"));
        assert_eq!(db.command.as_deref(), Some("Office Address List"));
        let x = v.to_xml(true).unwrap();
        assert_eq!(
            Connections::parse(&x).unwrap().connections[0]
                .database
                .as_ref()
                .unwrap()
                .command_type,
            Some(3)
        );
    }
    #[test]
    fn libreoffice_text_import_fields() {
        let v = f_without_broken_thumbnail(include_bytes!(
            "../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/queryTableExport.xlsx"
        ));
        assert_eq!(v.connections.len(), 2);
        assert_eq!(
            v.connections[0]
                .text
                .as_ref()
                .unwrap()
                .fields
                .as_ref()
                .unwrap()
                .len(),
            2
        );
        assert_eq!(v.connections[1].text.as_ref().unwrap().comma, Some(true));
    }
    #[test]
    fn libreoffice_olap_and_extensions() {
        let v = f(include_bytes!(
            "../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf66377.xlsx"
        ));
        assert_eq!(
            v.connections[0].olap.as_ref().unwrap().row_drill_count,
            Some(1000)
        );
        assert!(
            std::str::from_utf8(v.connections[1].extension_xml.as_deref().unwrap())
                .unwrap()
                .contains("x15:rangePr")
        );
    }
    #[test]
    fn libreoffice_prefixed_core_namespace() {
        let v = f(include_bytes!(
            "../../../test-data/libreoffice-core/sc/qa/unit/data/xlsx/tdf167689_xmlMaps_and_xmlColumnPr.xlsx"
        ));
        assert_eq!(
            v.connections[0].web.as_ref().unwrap().xml_source,
            Some(true)
        );
    }
    #[test]
    fn standards_parameters_tables_strict_and_mce() {
        let xml = format!(
            r#"<connections xmlns="{XS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><connection id="9" refreshedVersion="8" credentials="stored"><webPr htmlFormat="rtf"><tables count="3"><m/><s v="A"/><x v="2"/></tables></webPr><parameters count="1"><parameter name="p" sqlType="4" parameterType="value" double="1.5"/></parameters></connection></mc:Fallback></mc:AlternateContent></connections>"#
        );
        let v = Connections::parse(xml.as_bytes()).unwrap();
        assert_eq!(
            v.connections[0].parameters.as_ref().unwrap()[0].double,
            Some(1.5)
        );
        assert_eq!(Connections::parse(&v.to_xml(false).unwrap()).unwrap(), v);
    }
    #[test]
    fn rejects_malformed_and_unsafe() {
        for xml in [
            format!(r#"<connections xmlns="{X}"/>"#),
            format!(
                r#"<connections xmlns="{X}"><connection id="1" refreshedVersion="0"><parameters count="2"><parameter/></parameters></connection></connections>"#
            ),
            format!(
                r#"<connections xmlns="{X}"><connection id="1" refreshedVersion="0"><parameters><parameter double="NaN"/></parameters></connection></connections>"#
            ),
            format!(
                r#"<!DOCTYPE x><connections xmlns="{X}"><connection id="1" refreshedVersion="0"/></connections>"#
            ),
        ] {
            assert!(
                Connections::parse(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }
    fn package(content_type: &str, external: bool, outbound: bool) -> OpcPackage {
        let mut p = OpcPackage::new();
        let wb = PackURI::new("/xl/workbook.xml").unwrap();
        let mut w = BlobPart::new(
            wb,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            Vec::new(),
        );
        if external {
            w.rels_mut().add_relationship(
                REL.into(),
                "https://example.invalid/c.xml".into(),
                "rId1".into(),
                true,
            );
        } else {
            w.relate_to("connections.xml", REL);
        }
        p.relate_to(
            "xl/workbook.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );
        p.add_part(Box::new(w));
        let mut c=BlobPart::new(PackURI::new("/xl/connections.xml").unwrap(),content_type.into(),format!(r#"<connections xmlns="{X}"><connection id="1" refreshedVersion="0"/></connections>"#).into_bytes());
        if outbound {
            c.relate_to("other.xml", "urn:forbidden");
        }
        p.add_part(Box::new(c));
        p
    }
    #[test]
    fn rejects_external_wrong_content_and_outbound_package_edges() {
        assert!(load_from_package(&package(CT, true, false)).is_err());
        assert!(load_from_package(&package("application/xml", false, false)).is_err());
        assert!(load_from_package(&package(CT, false, true)).is_err());
    }
}
