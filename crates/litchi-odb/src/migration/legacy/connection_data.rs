//! Typed, inert ODF database connection metadata.

use super::document::{
    DatabaseContent, DatabaseDocument, DatabaseElement, DatabaseElementKind,
    parse_database_content, validate_database_root,
};
use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const DATABASE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const SCRIPT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
const MAX_XML_SIZE: usize = 64 * 1024 * 1024;
const MAX_VALUE_SIZE: usize = 1024 * 1024;
const MAX_REFERENCE_SIZE: usize = 64 * 1024;
const MAX_INTEGER_DIGITS: usize = 4096;

/// Canonical, arbitrary-width XML Schema `positiveInteger` metadata.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct OdfDatabasePositiveInteger(String);

impl OdfDatabasePositiveInteger {
    pub fn new(value: &str) -> Result<Self> {
        let value = value.trim_matches(|ch| matches!(ch, ' ' | '\t' | '\r' | '\n'));
        let digits = value.strip_prefix('+').unwrap_or(value);
        if digits.is_empty()
            || digits.len() > MAX_INTEGER_DIGITS
            || !digits.bytes().all(|byte| byte.is_ascii_digit())
            || digits.bytes().all(|byte| byte == b'0')
        {
            return Err(Error::InvalidFormat(
                "database positive integer has an invalid lexical value".to_string(),
            ));
        }
        let canonical = digits.trim_start_matches('0');
        Ok(Self(canonical.to_string()))
    }

    pub fn as_str(&self) -> &str {
        &self.0
    }
}

/// An inert `db:connection-resource`; its URI is never resolved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfOdbConnectionResource {
    pub href: String,
    pub show_none: bool,
    pub actuate_on_request: bool,
}

impl OdfOdbConnectionResource {
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            show_none: false,
            actuate_on_request: false,
        }
    }
}

/// A file-backed database description. The path remains inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseFileSource {
    pub href: String,
    pub media_type: String,
    pub extension: Option<String>,
}

/// The address branch selected by `db:server-database`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseServerLocation {
    Host {
        hostname: String,
        port: Option<OdfDatabasePositiveInteger>,
    },
    /// The RNG permits the local-socket attribute itself to be omitted.
    LocalSocket(Option<String>),
}

/// Inert server database metadata; no driver is loaded or contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseServerSource {
    pub database_type: String,
    pub location: OdfDatabaseServerLocation,
    pub database_name: Option<String>,
}

/// The exclusive source choice inside `db:connection-data`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseConnectionSource {
    Resource(OdfOdbConnectionResource),
    File(OdfDatabaseFileSource),
    Server(OdfDatabaseServerSource),
}

/// The optional login identity choice.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdfDatabaseLoginIdentity {
    UserName(String),
    UseSystemUser(bool),
}

/// Inert login policy. No credentials are requested or submitted.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct OdfDatabaseLogin {
    pub identity: Option<OdfDatabaseLoginIdentity>,
    pub is_password_required: Option<bool>,
    pub login_timeout: Option<OdfDatabasePositiveInteger>,
}

/// The complete typed content of one required `db:connection-data` element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OdfDatabaseConnectionData {
    pub source: OdfDatabaseConnectionSource,
    pub login: Option<OdfDatabaseLogin>,
}

impl OdfDatabaseConnectionData {
    pub fn to_xml_fragment(&self) -> Result<String> {
        validate_connection_data(self)?;
        let mut xml = format!(
            "<db:connection-data xmlns:db=\"{DATABASE_NAMESPACE}\" xmlns:xlink=\"{XLINK_NAMESPACE}\">"
        );
        push_source(&mut xml, &self.source);
        if let Some(login) = &self.login {
            push_login(&mut xml, login);
        }
        xml.push_str("</db:connection-data>");
        Ok(xml)
    }
}

impl DatabaseDocument {
    /// Returns the strictly typed connection metadata without activating it.
    pub fn connection_data(&self) -> Result<OdfDatabaseConnectionData> {
        connection_data_from_root(self.database())
    }
}

/// Strictly parses connection metadata from an ODF database `content.xml`.
pub fn parse_database_connection_data_xml(xml: &str) -> Result<OdfDatabaseConnectionData> {
    preflight_xml(xml)?;
    let database = parse_database_content(xml)?;
    validate_database_root(&database)?;
    connection_data_from_root(&database)
}

/// Losslessly replaces the required `db:connection-data` subtree.
pub fn replace_database_connection_data_xml(
    xml: &str,
    connection: &OdfDatabaseConnectionData,
) -> Result<String> {
    parse_database_connection_data_xml(xml)?;
    let (start, end) = locate_connection_data(xml)?;
    let replacement = mutation_fragment(connection)?;
    if !xml.is_char_boundary(start) || !xml.is_char_boundary(end) || start > end {
        return Err(Error::InvalidFormat(
            "invalid database mutation span".to_string(),
        ));
    }
    let mut output = String::with_capacity(xml.len() - (end - start) + replacement.len());
    output.push_str(&xml[..start]);
    output.push_str(&replacement);
    output.push_str(&xml[end..]);
    Ok(output)
}

pub(super) fn connection_data_from_root(
    root: &DatabaseElement,
) -> Result<OdfDatabaseConnectionData> {
    validate_database_root(root)?;
    let data_source = root
        .children()
        .find(|child| child.kind() == DatabaseElementKind::DataSource)
        .ok_or_else(|| Error::InvalidFormat("database has no db:data-source".to_string()))?;
    reject_active_descendants(data_source)?;
    require_no_attributes(data_source)?;
    let children = strict_children(data_source)?;
    if children.first().map(|child| child.kind()) != Some(DatabaseElementKind::ConnectionData) {
        return Err(Error::InvalidFormat(
            "db:data-source must begin with exactly one db:connection-data".to_string(),
        ));
    }
    let mut phase = 0u8;
    for (index, child) in children.iter().enumerate() {
        match child.kind() {
            DatabaseElementKind::ConnectionData if index == 0 => phase = 1,
            DatabaseElementKind::DriverSettings if phase == 1 => phase = 2,
            DatabaseElementKind::ApplicationConnectionSettings if phase == 1 || phase == 2 => {
                phase = 3
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "db:data-source children violate connection/driver/application order or cardinality"
                        .to_string(),
                ));
            },
        }
    }
    parse_connection_data(children[0])
}

fn parse_connection_data(element: &DatabaseElement) -> Result<OdfDatabaseConnectionData> {
    require_kind(element, DatabaseElementKind::ConnectionData)?;
    require_no_attributes(element)?;
    let children = strict_children(element)?;
    if !(1..=2).contains(&children.len()) {
        return Err(Error::InvalidFormat(
            "db:connection-data requires one source and optional login".to_string(),
        ));
    }
    let source = match children[0].kind() {
        DatabaseElementKind::ConnectionResource => {
            OdfDatabaseConnectionSource::Resource(parse_resource(children[0])?)
        },
        DatabaseElementKind::DatabaseDescription => parse_database_description(children[0])?,
        _ => {
            return Err(Error::InvalidFormat(
                "db:connection-data has an unsupported source child".to_string(),
            ));
        },
    };
    let login = if children.len() == 2 {
        if children[1].kind() != DatabaseElementKind::Login {
            return Err(Error::InvalidFormat(
                "only db:login may follow the connection source".to_string(),
            ));
        }
        Some(parse_login(children[1])?)
    } else {
        None
    };
    let value = OdfDatabaseConnectionData { source, login };
    validate_connection_data(&value)?;
    Ok(value)
}

fn parse_database_description(element: &DatabaseElement) -> Result<OdfDatabaseConnectionSource> {
    require_no_attributes(element)?;
    let children = strict_children(element)?;
    if children.len() != 1 {
        return Err(Error::InvalidFormat(
            "db:database-description requires exactly one database description".to_string(),
        ));
    }
    match children[0].kind() {
        DatabaseElementKind::FileBasedDatabase => Ok(OdfDatabaseConnectionSource::File(
            parse_file_source(children[0])?,
        )),
        DatabaseElementKind::ServerDatabase => Ok(OdfDatabaseConnectionSource::Server(
            parse_server_source(children[0])?,
        )),
        _ => Err(Error::InvalidFormat(
            "db:database-description accepts only file or server databases".to_string(),
        )),
    }
}

fn parse_resource(element: &DatabaseElement) -> Result<OdfOdbConnectionResource> {
    require_empty(element)?;
    allow_attributes(
        element,
        &[
            (XLINK_NAMESPACE, "type"),
            (XLINK_NAMESPACE, "href"),
            (XLINK_NAMESPACE, "show"),
            (XLINK_NAMESPACE, "actuate"),
        ],
    )?;
    require_exact_attribute(element, XLINK_NAMESPACE, "type", "simple")?;
    let href = required_attribute(element, XLINK_NAMESPACE, "href")?.to_string();
    let show_none = optional_fixed_attribute(element, XLINK_NAMESPACE, "show", "none")?;
    let actuate_on_request =
        optional_fixed_attribute(element, XLINK_NAMESPACE, "actuate", "onRequest")?;
    Ok(OdfOdbConnectionResource {
        href,
        show_none,
        actuate_on_request,
    })
}

fn parse_file_source(element: &DatabaseElement) -> Result<OdfDatabaseFileSource> {
    require_empty(element)?;
    allow_attributes(
        element,
        &[
            (XLINK_NAMESPACE, "type"),
            (XLINK_NAMESPACE, "href"),
            (DATABASE_NAMESPACE, "media-type"),
            (DATABASE_NAMESPACE, "extension"),
        ],
    )?;
    require_exact_attribute(element, XLINK_NAMESPACE, "type", "simple")?;
    Ok(OdfDatabaseFileSource {
        href: required_attribute(element, XLINK_NAMESPACE, "href")?.to_string(),
        media_type: required_attribute(element, DATABASE_NAMESPACE, "media-type")?.to_string(),
        extension: element
            .attribute(Some(DATABASE_NAMESPACE), "extension")
            .map(str::to_string),
    })
}

fn parse_server_source(element: &DatabaseElement) -> Result<OdfDatabaseServerSource> {
    require_empty(element)?;
    allow_attributes(
        element,
        &[
            (DATABASE_NAMESPACE, "type"),
            (DATABASE_NAMESPACE, "hostname"),
            (DATABASE_NAMESPACE, "port"),
            (DATABASE_NAMESPACE, "local-socket"),
            (DATABASE_NAMESPACE, "database-name"),
        ],
    )?;
    let database_type = required_attribute(element, DATABASE_NAMESPACE, "type")?.to_string();
    validate_qname(&database_type)?;
    let hostname = element.attribute(Some(DATABASE_NAMESPACE), "hostname");
    let port = element.attribute(Some(DATABASE_NAMESPACE), "port");
    let local_socket = element.attribute(Some(DATABASE_NAMESPACE), "local-socket");
    let location = if let Some(hostname) = hostname {
        if local_socket.is_some() {
            return Err(Error::InvalidFormat(
                "db:server-database cannot combine hostname and local-socket".to_string(),
            ));
        }
        OdfDatabaseServerLocation::Host {
            hostname: hostname.to_string(),
            port: port.map(OdfDatabasePositiveInteger::new).transpose()?,
        }
    } else {
        if port.is_some() {
            return Err(Error::InvalidFormat(
                "db:port requires db:hostname".to_string(),
            ));
        }
        OdfDatabaseServerLocation::LocalSocket(local_socket.map(str::to_string))
    };
    Ok(OdfDatabaseServerSource {
        database_type,
        location,
        database_name: element
            .attribute(Some(DATABASE_NAMESPACE), "database-name")
            .map(str::to_string),
    })
}

fn parse_login(element: &DatabaseElement) -> Result<OdfDatabaseLogin> {
    require_empty(element)?;
    allow_attributes(
        element,
        &[
            (DATABASE_NAMESPACE, "user-name"),
            (DATABASE_NAMESPACE, "use-system-user"),
            (DATABASE_NAMESPACE, "is-password-required"),
            (DATABASE_NAMESPACE, "login-timeout"),
        ],
    )?;
    let user_name = element.attribute(Some(DATABASE_NAMESPACE), "user-name");
    let system_user = element.attribute(Some(DATABASE_NAMESPACE), "use-system-user");
    if user_name.is_some() && system_user.is_some() {
        return Err(Error::InvalidFormat(
            "db:login identity attributes are mutually exclusive".to_string(),
        ));
    }
    let identity = if let Some(value) = user_name {
        Some(OdfDatabaseLoginIdentity::UserName(value.to_string()))
    } else if let Some(value) = system_user {
        Some(OdfDatabaseLoginIdentity::UseSystemUser(parse_boolean(
            value,
        )?))
    } else {
        None
    };
    Ok(OdfDatabaseLogin {
        identity,
        is_password_required: element
            .attribute(Some(DATABASE_NAMESPACE), "is-password-required")
            .map(parse_boolean)
            .transpose()?,
        login_timeout: element
            .attribute(Some(DATABASE_NAMESPACE), "login-timeout")
            .map(OdfDatabasePositiveInteger::new)
            .transpose()?,
    })
}

fn strict_children(element: &DatabaseElement) -> Result<Vec<&DatabaseElement>> {
    let mut children = Vec::new();
    for content in element.content() {
        match content {
            DatabaseContent::Text(text) if text.trim().is_empty() => {},
            DatabaseContent::Text(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{} must not contain character data",
                    element.local_name()
                )));
            },
            DatabaseContent::Element(child) => children.push(child),
        }
    }
    Ok(children)
}

fn require_empty(element: &DatabaseElement) -> Result<()> {
    if strict_children(element)?.is_empty() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{} must be empty",
            element.local_name()
        )))
    }
}

fn require_kind(element: &DatabaseElement, kind: DatabaseElementKind) -> Result<()> {
    if element.kind() == kind {
        Ok(())
    } else {
        Err(Error::InvalidFormat(
            "wrong database element kind".to_string(),
        ))
    }
}

fn require_no_attributes(element: &DatabaseElement) -> Result<()> {
    if element.attributes().is_empty() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{} does not accept attributes",
            element.local_name()
        )))
    }
}

fn allow_attributes(element: &DatabaseElement, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in element.attributes() {
        if !allowed.iter().any(|(namespace, local)| {
            attribute.namespace_uri() == Some(*namespace) && attribute.local_name() == *local
        }) {
            return Err(Error::InvalidFormat(format!(
                "unsupported {} attribute {}",
                element.local_name(),
                attribute.local_name()
            )));
        }
    }
    Ok(())
}

fn required_attribute<'a>(
    element: &'a DatabaseElement,
    namespace: &str,
    local: &str,
) -> Result<&'a str> {
    element
        .attribute(Some(namespace), local)
        .ok_or_else(|| Error::InvalidFormat(format!("{} requires {local}", element.local_name())))
}

fn require_exact_attribute(
    element: &DatabaseElement,
    namespace: &str,
    local: &str,
    expected: &str,
) -> Result<()> {
    if required_attribute(element, namespace, local)? == expected {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "{} {local} must be {expected}",
            element.local_name()
        )))
    }
}

fn optional_fixed_attribute(
    element: &DatabaseElement,
    namespace: &str,
    local: &str,
    expected: &str,
) -> Result<bool> {
    match element.attribute(Some(namespace), local) {
        None => Ok(false),
        Some(value) if value == expected => Ok(true),
        Some(_) => Err(Error::InvalidFormat(format!(
            "{} {local} must be {expected}",
            element.local_name()
        ))),
    }
}

fn parse_boolean(value: &str) -> Result<bool> {
    match value {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(Error::InvalidFormat(
            "database boolean must be true or false".to_string(),
        )),
    }
}

fn validate_connection_data(connection: &OdfDatabaseConnectionData) -> Result<()> {
    match &connection.source {
        OdfDatabaseConnectionSource::Resource(resource) => {
            validate_value(&resource.href, "connection URI", MAX_REFERENCE_SIZE)?;
        },
        OdfDatabaseConnectionSource::File(file) => {
            validate_value(&file.href, "database file URI", MAX_REFERENCE_SIZE)?;
            validate_value(&file.media_type, "database media type", MAX_VALUE_SIZE)?;
            if let Some(extension) = file.extension.as_deref() {
                validate_value(extension, "database extension", MAX_VALUE_SIZE)?;
            }
        },
        OdfDatabaseConnectionSource::Server(server) => {
            validate_value(&server.database_type, "database type", MAX_VALUE_SIZE)?;
            validate_qname(&server.database_type)?;
            match &server.location {
                OdfDatabaseServerLocation::Host { hostname, .. } => {
                    validate_value(hostname, "database hostname", MAX_VALUE_SIZE)?;
                },
                OdfDatabaseServerLocation::LocalSocket(socket) => {
                    if let Some(socket) = socket.as_deref() {
                        validate_value(socket, "database local socket", MAX_VALUE_SIZE)?;
                    }
                },
            }
            if let Some(name) = server.database_name.as_deref() {
                validate_value(name, "database name", MAX_VALUE_SIZE)?;
            }
        },
    }
    if let Some(login) = &connection.login
        && let Some(OdfDatabaseLoginIdentity::UserName(user)) = login.identity.as_ref()
    {
        validate_value(user, "database user name", MAX_VALUE_SIZE)?;
    }
    Ok(())
}

fn validate_qname(value: &str) -> Result<()> {
    let mut parts = value.split(':');
    let first = parts.next().unwrap_or_default();
    let second = parts.next();
    if parts.next().is_some()
        || !valid_ncname(first)
        || second.is_some_and(|part| !valid_ncname(part))
    {
        return Err(Error::InvalidFormat(
            "db:type must be an XML QName lexical value".to_string(),
        ));
    }
    Ok(())
}

fn valid_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    chars
        .next()
        .is_some_and(|ch| ch == '_' || ch.is_alphabetic())
        && chars.all(|ch| ch == '_' || ch == '-' || ch == '.' || ch.is_alphanumeric())
}

fn validate_value(value: &str, name: &str, limit: usize) -> Result<()> {
    if value.len() > limit {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds {limit} bytes"
        )));
    }
    if !value.chars().all(is_xml_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} contains forbidden XML characters"
        )));
    }
    Ok(())
}

fn is_xml_char(ch: char) -> bool {
    matches!(ch, '\u{9}' | '\u{A}' | '\u{D}')
        || ('\u{20}'..='\u{D7FF}').contains(&ch)
        || ('\u{E000}'..='\u{FFFD}').contains(&ch)
        || ('\u{10000}'..='\u{10FFFF}').contains(&ch)
}

fn reject_active_descendants(element: &DatabaseElement) -> Result<()> {
    if (element.namespace_uri() == Some(OFFICE_NAMESPACE)
        && element.local_name() == "event-listeners")
        || element.namespace_uri() == Some(SCRIPT_NAMESPACE)
    {
        return Err(Error::InvalidFormat(
            "active behavior is not accepted in database connection metadata".to_string(),
        ));
    }
    for child in element.children() {
        reject_active_descendants(child)?;
    }
    Ok(())
}

fn push_source(xml: &mut String, source: &OdfDatabaseConnectionSource) {
    match source {
        OdfDatabaseConnectionSource::Resource(resource) => {
            xml.push_str("<db:connection-resource xlink:type=\"simple\" xlink:href=\"");
            push_attribute(xml, &resource.href);
            xml.push('"');
            if resource.show_none {
                xml.push_str(" xlink:show=\"none\"");
            }
            if resource.actuate_on_request {
                xml.push_str(" xlink:actuate=\"onRequest\"");
            }
            xml.push_str("/>");
        },
        OdfDatabaseConnectionSource::File(file) => {
            xml.push_str("<db:database-description><db:file-based-database xlink:type=\"simple\" xlink:href=\"");
            push_attribute(xml, &file.href);
            xml.push_str("\" db:media-type=\"");
            push_attribute(xml, &file.media_type);
            xml.push('"');
            if let Some(extension) = file.extension.as_deref() {
                xml.push_str(" db:extension=\"");
                push_attribute(xml, extension);
                xml.push('"');
            }
            xml.push_str("/></db:database-description>");
        },
        OdfDatabaseConnectionSource::Server(server) => {
            xml.push_str("<db:database-description><db:server-database db:type=\"");
            push_attribute(xml, &server.database_type);
            xml.push('"');
            match &server.location {
                OdfDatabaseServerLocation::Host { hostname, port } => {
                    xml.push_str(" db:hostname=\"");
                    push_attribute(xml, hostname);
                    xml.push('"');
                    if let Some(port) = port {
                        xml.push_str(" db:port=\"");
                        xml.push_str(port.as_str());
                        xml.push('"');
                    }
                },
                OdfDatabaseServerLocation::LocalSocket(socket) => {
                    if let Some(socket) = socket.as_deref() {
                        xml.push_str(" db:local-socket=\"");
                        push_attribute(xml, socket);
                        xml.push('"');
                    }
                },
            }
            if let Some(name) = server.database_name.as_deref() {
                xml.push_str(" db:database-name=\"");
                push_attribute(xml, name);
                xml.push('"');
            }
            xml.push_str("/></db:database-description>");
        },
    }
}

fn push_login(xml: &mut String, login: &OdfDatabaseLogin) {
    xml.push_str("<db:login");
    match login.identity.as_ref() {
        Some(OdfDatabaseLoginIdentity::UserName(value)) => {
            xml.push_str(" db:user-name=\"");
            push_attribute(xml, value);
            xml.push('"');
        },
        Some(OdfDatabaseLoginIdentity::UseSystemUser(value)) => {
            xml.push_str(" db:use-system-user=\"");
            xml.push_str(if *value { "true" } else { "false" });
            xml.push('"');
        },
        None => {},
    }
    if let Some(value) = login.is_password_required {
        xml.push_str(" db:is-password-required=\"");
        xml.push_str(if value { "true" } else { "false" });
        xml.push('"');
    }
    if let Some(value) = &login.login_timeout {
        xml.push_str(" db:login-timeout=\"");
        xml.push_str(value.as_str());
        xml.push('"');
    }
    xml.push_str("/>");
}

fn push_attribute(xml: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => xml.push_str("&amp;"),
            '<' => xml.push_str("&lt;"),
            '"' => xml.push_str("&quot;"),
            '\r' => xml.push_str("&#13;"),
            '\n' => xml.push_str("&#10;"),
            '\t' => xml.push_str("&#9;"),
            _ => xml.push(ch),
        }
    }
}

fn mutation_fragment(connection: &OdfDatabaseConnectionData) -> Result<String> {
    connection.to_xml_fragment()
}

fn preflight_xml(xml: &str) -> Result<()> {
    if xml.len() > MAX_XML_SIZE {
        return Err(Error::InvalidFormat(format!(
            "database XML exceeds {MAX_XML_SIZE} bytes"
        )));
    }
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid database XML: {error}")))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let namespace = namespace_string(&namespace)?;
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 database element name".to_string())
                })?;
                if (namespace.as_deref() == Some(OFFICE_NAMESPACE) && local == "event-listeners")
                    || namespace.as_deref() == Some(SCRIPT_NAMESPACE)
                {
                    return Err(Error::InvalidFormat(
                        "active behavior is not accepted in database connection XML".to_string(),
                    ));
                }
            },
            Event::DocType(_) | Event::PI(_) | Event::GeneralRef(_) => {
                return Err(Error::InvalidFormat(
                    "DTD, processing instructions, and entity references are not accepted in database connection XML"
                        .to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(())
}

fn locate_connection_data(xml: &str) -> Result<(usize, usize)> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut active: Option<(usize, usize)> = None;
    loop {
        let start = reader.buffer_position() as usize;
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid database XML: {error}")))?;
        let namespace = namespace_string(&namespace)?;
        let end = reader.buffer_position() as usize;
        match event {
            Event::Start(ref element) => {
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 database element name".to_string())
                })?;
                if namespace.as_deref() == Some(DATABASE_NAMESPACE) && local == "connection-data" {
                    if active.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested db:connection-data is invalid".to_string(),
                        ));
                    }
                    active = Some((depth, start));
                }
                depth += 1;
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("database XML depth underflow".to_string())
                })?;
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 database element name".to_string())
                })?;
                if namespace.as_deref() == Some(DATABASE_NAMESPACE)
                    && local == "connection-data"
                    && active.is_some_and(|(active_depth, _)| active_depth == depth)
                {
                    let (_, span_start) = active.take().expect("checked above");
                    return Ok((span_start, end));
                }
            },
            Event::Empty(ref element) => {
                let local_name = element.local_name();
                let local = std::str::from_utf8(local_name.as_ref()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 database element name".to_string())
                })?;
                if namespace.as_deref() == Some(DATABASE_NAMESPACE) && local == "connection-data" {
                    return Ok((start, end));
                }
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Err(Error::InvalidFormat(
        "database XML has no db:connection-data".to_string(),
    ))
}

fn namespace_string(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(value)) => std::str::from_utf8(value)
            .map(|value| Some(value.to_string()))
            .map_err(|_| Error::InvalidFormat("non-UTF-8 namespace URI".to_string())),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown database namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;

    fn wrap(connection: &str, tail: &str) -> String {
        format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DATABASE_NAMESPACE}" xmlns:x="{XLINK_NAMESPACE}"><o:body><o:database><d:data-source>{connection}{tail}</d:data-source></o:database></o:body></o:document-content>"#
        )
    }

    #[test]
    fn parses_and_roundtrips_every_source_and_login_choice_inertly() {
        let documents = [
            wrap(
                r#"<d:connection-data><d:connection-resource x:type="simple" x:href="sdbc:embedded:firebird" x:show="none" x:actuate="onRequest"/><d:login d:user-name="alice" d:is-password-required="true" d:login-timeout="0007"/></d:connection-data>"#,
                "",
            ),
            wrap(
                r#"<d:connection-data><d:database-description><d:file-based-database x:type="simple" x:href="file:///tmp/a.db" d:media-type="application/x-db" d:extension="db"/></d:database-description><d:login d:use-system-user="false"/></d:connection-data>"#,
                "<d:driver-settings/><d:application-connection-settings/>",
            ),
            wrap(
                r#"<d:connection-data><d:database-description><d:server-database d:type="mysql" d:hostname="db.example" d:port="3306" d:database-name="main"/></d:database-description></d:connection-data>"#,
                "",
            ),
            wrap(
                r#"<d:connection-data><d:database-description><d:server-database d:type="postgres" d:local-socket="/run/postgresql"/></d:database-description></d:connection-data>"#,
                "",
            ),
        ];
        for xml in documents {
            let parsed = parse_database_connection_data_xml(&xml).unwrap();
            let canonical = parsed.to_xml_fragment().unwrap();
            let reparsed = parse_database_connection_data_xml(&wrap(&canonical, "")).unwrap();
            assert_eq!(reparsed, parsed);
        }
    }

    #[test]
    fn rejects_schema_violations_active_content_and_unbounded_lexicals() {
        let invalid = [
            r#"<d:connection-data/>"#,
            r#"<d:connection-data bad="x"><d:connection-resource x:type="simple" x:href="db"/></d:connection-data>"#,
            r#"<d:connection-data><d:connection-resource x:href="db"/></d:connection-data>"#,
            r#"<d:connection-data><d:connection-resource x:type="extended" x:href="db"/></d:connection-data>"#,
            r#"<d:connection-data><d:connection-resource x:type="simple" x:href="db"/><d:login d:user-name="a" d:use-system-user="true"/></d:connection-data>"#,
            r#"<d:connection-data><d:database-description><d:server-database d:type="mysql" d:port="3"/></d:database-description></d:connection-data>"#,
            r#"<d:connection-data><d:database-description><d:server-database d:type="bad:type:token"/></d:database-description></d:connection-data>"#,
            r#"<d:connection-data><d:database-description><d:file-based-database x:type="simple" x:href="db"/></d:database-description></d:connection-data>"#,
            r#"<d:connection-data><o:event-listeners/></d:connection-data>"#,
        ];
        for connection in invalid {
            let xml = wrap(connection, "");
            assert!(
                parse_database_connection_data_xml(&xml).is_err(),
                "accepted {connection}"
            );
        }
        let wrong_order = wrap(
            r#"<d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data>"#,
            "<d:application-connection-settings/><d:driver-settings/>",
        );
        assert!(parse_database_connection_data_xml(&wrong_order).is_err());
        let doctype = format!(
            "<!DOCTYPE x [<!ENTITY e 'x'>]>{}",
            wrap(
                r#"<d:connection-data><d:connection-resource x:type="simple" x:href="db"/></d:connection-data>"#,
                ""
            )
        );
        assert!(parse_database_connection_data_xml(&doctype).is_err());
        assert!(OdfDatabasePositiveInteger::new("0").is_err());
        assert_eq!(
            OdfDatabasePositiveInteger::new("+00042").unwrap().as_str(),
            "42"
        );
    }

    #[test]
    fn lossless_replacement_and_database_document_api_preserve_unrelated_content() {
        let original = wrap(
            r#"<d:connection-data><d:connection-resource x:type="simple" x:href="old"/></d:connection-data>"#,
            r#"<d:driver-settings d:parameter-name-substitution="false"/>"#,
        );
        let replacement = OdfDatabaseConnectionData {
            source: OdfDatabaseConnectionSource::Server(OdfDatabaseServerSource {
                database_type: "postgres".into(),
                location: OdfDatabaseServerLocation::Host {
                    hostname: "localhost".into(),
                    port: Some(OdfDatabasePositiveInteger::new("5432").unwrap()),
                },
                database_name: Some("inventory".into()),
            }),
            login: Some(OdfDatabaseLogin {
                identity: Some(OdfDatabaseLoginIdentity::UseSystemUser(true)),
                is_password_required: Some(false),
                login_timeout: None,
            }),
        };
        let replaced = replace_database_connection_data_xml(&original, &replacement).unwrap();
        assert!(replaced.contains("parameter-name-substitution"));
        assert_eq!(
            parse_database_connection_data_xml(&replaced).unwrap(),
            replacement
        );

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_DATABASE).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, replaced.as_bytes())
            .unwrap();
        let document = DatabaseDocument::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        assert_eq!(document.connection_data().unwrap(), replacement);
    }
}
