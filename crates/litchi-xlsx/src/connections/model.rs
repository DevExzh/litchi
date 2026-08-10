//! Typed `SpreadsheetML` external connection definitions and collection editing.

use super::codec::bounded;
use super::{codec, invalid};
use litchi_core::sheet::Result;
use std::collections::{HashMap, HashSet};

// XML namespaces and package limits are kept with the typed owner so every layer
// uses the same contextual bounds.
pub(super) const CORE_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(super) const STRICT_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(super) const CONNECTIONS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
pub(super) const STRICT_CONNECTIONS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";
pub(super) const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
pub(super) const QUERY_TABLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
pub(super) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(super) const MAX_EXTENSION_BYTES: usize = 8 * 1024 * 1024;
pub(super) const MAX_DOM_DEPTH: usize = 256;
pub(super) const MAX_DOM_NODES: usize = 1_000_000;
pub(super) const MAX_CONNECTIONS: usize = 65_536;
pub(super) const MAX_PARAMETERS: usize = 65_536;
pub(super) const MAX_WEB_TABLES: usize = 65_536;
pub(super) const MAX_TEXT_FIELDS: usize = 2001;
pub(super) const MAX_STRING_BYTES: usize = 1024 * 1024;

/// `SpreadsheetML` namespace conformance used when authoring a connection part.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn strict(self) -> bool {
        matches!(self, Self::Strict)
    }
}

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

#[allow(
    clippy::missing_fields_in_debug,
    reason = "sensitive text-import fields are deliberately omitted or redacted"
)]
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

#[allow(
    clippy::missing_fields_in_debug,
    reason = "all parameter payload alternatives are intentionally represented by one redacted value field"
)]
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
    #[must_use]
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
        validate_connections(self.connections.iter().chain(std::iter::once(&connection)))?;
        self.connections
            .try_reserve(1)
            .map_err(|_source| invalid("connection collection allocation failed"))?;
        self.connections.push(connection);
        Ok(())
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
        validate_connections(
            self.connections
                .iter()
                .enumerate()
                .map(|(index, candidate)| {
                    if index == offset {
                        &connection
                    } else {
                        candidate
                    }
                }),
        )?;
        self.connections[offset] = connection;
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
        validate_connections(
            self.connections
                .iter()
                .enumerate()
                .filter_map(|(index, connection)| (index != offset).then_some(connection)),
        )?;
        self.connections.remove(offset);
        Ok(true)
    }

    pub fn reorder(&mut self, ordered_ids: &[u32]) -> Result<()> {
        if ordered_ids.len() != self.connections.len() {
            return Err(invalid("connection reorder must contain every ID"));
        }
        let mut index_by_id = HashMap::new();
        index_by_id
            .try_reserve(self.connections.len())
            .map_err(|_source| invalid("connection reorder allocation failed"))?;
        for (index, connection) in self.connections.iter().enumerate() {
            if index_by_id.insert(connection.id, index).is_some() {
                return Err(invalid("connection reorder is not a permutation"));
            }
        }

        let mut seen = HashSet::new();
        seen.try_reserve(ordered_ids.len())
            .map_err(|_source| invalid("connection reorder allocation failed"))?;
        let mut ranks = Vec::new();
        ranks
            .try_reserve_exact(ordered_ids.len())
            .map_err(|_source| invalid("connection reorder allocation failed"))?;
        ranks.resize(ordered_ids.len(), usize::MAX);
        for (rank, id) in ordered_ids.iter().enumerate() {
            let Some(&index) = index_by_id.get(id) else {
                return Err(invalid("connection reorder is not a permutation"));
            };
            if !seen.insert(*id) {
                return Err(invalid("connection reorder is not a permutation"));
            }
            ranks[index] = rank;
        }
        self.connections.sort_unstable_by_key(|connection| {
            index_by_id
                .get(&connection.id)
                .and_then(|index| ranks.get(*index))
                .copied()
                .unwrap_or(usize::MAX)
        });
        Ok(())
    }
}

/// Store the complete inert connection set and validate every query-table reference first.
pub(super) fn validate(v: &Connections) -> Result<()> {
    validate_connections(v.connections.iter())
}

fn validate_connections<'a, I>(connections: I) -> Result<()>
where
    I: IntoIterator<Item = &'a Connection>,
{
    let connections = connections.into_iter();
    let (lower, upper) = connections.size_hint();
    let reserve = upper.unwrap_or(lower).min(MAX_CONNECTIONS);
    let mut ids = HashSet::new();
    let mut names = HashSet::new();
    ids.try_reserve(reserve)
        .map_err(|_source| invalid("connection validation allocation failed"))?;
    names
        .try_reserve(reserve)
        .map_err(|_source| invalid("connection validation allocation failed"))?;
    let mut count = 0usize;
    let mut ext = 0usize;
    for c in connections {
        count = count
            .checked_add(1)
            .ok_or_else(|| invalid("connection count overflow"))?;
        if count > MAX_CONNECTIONS {
            return Err(invalid("invalid connection count"));
        }
        ids.try_reserve(1)
            .map_err(|_source| invalid("connection validation allocation failed"))?;
        if !ids.insert(c.id) {
            return Err(invalid("duplicate connection id"));
        }
        if let Some(n) = &c.name {
            bounded(n)?;
            names
                .try_reserve(1)
                .map_err(|_source| invalid("connection validation allocation failed"))?;
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
            .is_some_and(|x| x.is_empty() || x.len() > MAX_WEB_TABLES)
        {
            return Err(invalid("invalid web table count"));
        }
        if c.text
            .as_ref()
            .and_then(|x| x.fields.as_ref())
            .is_some_and(|x| x.is_empty() || x.len() > MAX_TEXT_FIELDS)
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
            codec::parse_dom(e)?;
            ext = ext
                .checked_add(e.len())
                .ok_or_else(|| invalid("extension size overflow"))?;
        }
    }
    if count == 0 {
        return Err(invalid("invalid connection count"));
    }
    if ext > MAX_EXTENSION_BYTES {
        return Err(invalid("connection extensions exceed 8 MiB"));
    }
    Ok(())
}
