//! Bounded, inert ODB schema and query catalogs.

use super::{
    query::Query,
    table::{Column, Table, TableKind},
};
use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";

/// Finite limits for semantic ODB catalog discovery.
#[allow(
    clippy::struct_field_names,
    reason = "each field names its stable public max-setting builder counterpart"
)]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_xml_bytes: usize,
    max_events: usize,
    max_depth: usize,
    max_tables: usize,
    max_columns: usize,
    max_queries: usize,
    max_attribute_bytes: usize,
}

impl Limits {
    /// Sets the maximum accepted `content.xml` byte length.
    #[must_use]
    pub const fn with_max_xml_bytes(mut self, value: usize) -> Self {
        self.max_xml_bytes = value;
        self
    }

    /// Sets the maximum XML event count.
    #[must_use]
    pub const fn with_max_events(mut self, value: usize) -> Self {
        self.max_events = value;
        self
    }

    /// Sets the maximum nested element depth.
    #[must_use]
    pub const fn with_max_depth(mut self, value: usize) -> Self {
        self.max_depth = value;
        self
    }

    /// Sets the maximum combined table declarations.
    #[must_use]
    pub const fn with_max_tables(mut self, value: usize) -> Self {
        self.max_tables = value;
        self
    }

    /// Sets the maximum combined column declarations.
    #[must_use]
    pub const fn with_max_columns(mut self, value: usize) -> Self {
        self.max_columns = value;
        self
    }

    /// Sets the maximum query declarations.
    #[must_use]
    pub const fn with_max_queries(mut self, value: usize) -> Self {
        self.max_queries = value;
        self
    }

    /// Sets the maximum encoded length of one semantic attribute.
    #[must_use]
    pub const fn with_max_attribute_bytes(mut self, value: usize) -> Self {
        self.max_attribute_bytes = value;
        self
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_xml_bytes: 64 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 512,
            max_tables: 65_536,
            max_columns: 1_000_000,
            max_queries: 65_536,
            max_attribute_bytes: 1024 * 1024,
        }
    }
}

/// A read-only catalog tied to the source package that produced it.
///
/// The catalog owns only decoded semantic strings. The full source XML and all
/// unknown markup remain borrowed through its source package and are never
/// rewritten or interpreted as executable content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Catalog<'source> {
    source: &'source str,
    owned: OwnedCatalog,
}

impl<'source> Catalog<'source> {
    pub(crate) fn parse(source: &'source str, limits: Limits) -> Result<Self> {
        Ok(Self {
            source,
            owned: parse(source, limits)?,
        })
    }

    /// Returns the original content part that backs this read-only catalog.
    #[must_use]
    pub const fn source_xml(&self) -> &'source str {
        self.source
    }

    /// Returns table declarations in source order.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        self.owned.tables()
    }

    /// Returns stored queries in source order.
    #[must_use]
    pub fn queries(&self) -> &[Query] {
        self.owned.queries()
    }

    /// Finds one unambiguous table declaration by exact producer-visible name.
    ///
    /// # Errors
    ///
    /// Returns an error when the source contains more than one declaration
    /// with the requested name.
    pub fn table(&self, name: &str) -> Result<Option<&Table>> {
        select(self.tables(), name, Table::name, "table")
    }

    /// Finds one unambiguous stored query by exact producer-visible name.
    ///
    /// # Errors
    ///
    /// Returns an error when the source contains more than one query with the
    /// requested name.
    pub fn query(&self, name: &str) -> Result<Option<&Query>> {
        select(self.queries(), name, Query::name, "query")
    }

    /// Clones this source-bound read view into a detached semantic catalog.
    #[must_use]
    pub fn to_owned(&self) -> OwnedCatalog {
        self.owned.clone()
    }
}

/// A detached inert ODB semantic catalog.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct OwnedCatalog {
    tables: Vec<Table>,
    queries: Vec<Query>,
}

impl OwnedCatalog {
    /// Returns table declarations in source order.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Returns stored queries in source order.
    #[must_use]
    pub fn queries(&self) -> &[Query] {
        &self.queries
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Element {
    Document,
    Body,
    Database,
    DataSource,
    Queries,
    QueryCollection,
    TableRepresentation,
    TableRepresentations,
    TableDefinition,
    TableDefinitions,
    SchemaDefinition,
    Columns,
    ColumnDefinitions,
    Column,
    ColumnDefinition,
    Query,
    Other,
}

#[derive(Clone, Copy)]
struct Frame {
    element: Element,
    in_database: bool,
    table: Option<usize>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Database,
    Other,
}

fn parse(source: &str, limits: Limits) -> Result<OwnedCatalog> {
    if source.len() > limits.max_xml_bytes {
        return Err(invalid(
            "ODB semantic catalog source exceeds the byte limit",
        ));
    }

    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut stack = Vec::<Frame>::new();
    let mut catalog = OwnedCatalog::default();
    let mut events = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut database_seen = false;
    let mut data_sources = 0usize;
    let mut columns = 0usize;

    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| invalid("ODB semantic catalog event count overflow"))?;
        if events > limits.max_events {
            return Err(invalid("ODB semantic catalog exceeds the event limit"));
        }

        let (resolved, raw_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| invalid(&format!("invalid ODB semantic XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let event = raw_event.into_owned();
        match event {
            Event::Start(element) => {
                let frame = start(
                    &reader,
                    namespace,
                    &element,
                    stack.last().copied(),
                    &mut catalog,
                    &mut database_seen,
                    &mut data_sources,
                    &mut columns,
                    limits,
                )?;
                if stack.is_empty() {
                    if root_seen || root_closed || frame.element != Element::Document {
                        return Err(invalid(
                            "ODB semantic catalog has no office:document-content root",
                        ));
                    }
                    root_seen = true;
                }
                if stack.len() >= limits.max_depth {
                    return Err(invalid("ODB semantic catalog exceeds the nesting limit"));
                }
                stack.push(frame);
            },
            Event::Empty(element) => {
                if stack.is_empty() {
                    return Err(invalid("ODB semantic catalog root cannot be empty"));
                }
                if stack.len() >= limits.max_depth {
                    return Err(invalid("ODB semantic catalog exceeds the nesting limit"));
                }
                let _frame = start(
                    &reader,
                    namespace,
                    &element,
                    stack.last().copied(),
                    &mut catalog,
                    &mut database_seen,
                    &mut data_sources,
                    &mut columns,
                    limits,
                )?;
            },
            Event::End(_) => {
                if stack.pop().is_none() {
                    return Err(invalid("ODB semantic catalog has an unmatched closing tag"));
                }
                if stack.is_empty() {
                    root_closed = true;
                }
            },
            Event::DocType(_) => {
                return Err(invalid("DOCTYPE is not permitted in ODB content.xml"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if !root_seen || !root_closed || !stack.is_empty() {
        return Err(invalid("ODB semantic catalog XML is incomplete"));
    }
    if !database_seen || data_sources != 1 {
        return Err(invalid(
            "ODB semantic catalog has no valid office:database body",
        ));
    }
    Ok(catalog)
}

#[allow(
    clippy::too_many_arguments,
    reason = "one XML event updates the bounded catalog state"
)]
fn start(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    parent: Option<Frame>,
    catalog: &mut OwnedCatalog,
    database_seen: &mut bool,
    data_sources: &mut usize,
    columns: &mut usize,
    limits: Limits,
) -> Result<Frame> {
    let local = element.local_name();
    let kind = classify(parent, namespace, local.as_ref());
    let in_database = parent.is_some_and(|frame| frame.in_database) || kind == Element::Database;
    if kind == Element::Other && is_catalog_node(namespace, local.as_ref()) {
        return Err(invalid("ODB schema node has an invalid parent"));
    }
    let mut table = parent.and_then(|frame| frame.table);

    if kind == Element::Database {
        if *database_seen {
            return Err(invalid(
                "ODB semantic catalog has multiple office:database bodies",
            ));
        }
        *database_seen = true;
    }
    if kind == Element::DataSource {
        *data_sources = data_sources
            .checked_add(1)
            .ok_or_else(|| invalid("ODB semantic catalog data-source count overflow"))?;
    }
    if in_database
        && matches!(
            kind,
            Element::TableRepresentation | Element::TableDefinition
        )
    {
        ensure_capacity(
            catalog.tables.len(),
            limits.max_tables,
            "table declarations",
        )?;
        let name = required_db_attr(reader, element, b"name", limits)?;
        let table_kind = if kind == Element::TableRepresentation {
            TableKind::Representation
        } else {
            TableKind::Definition
        };
        catalog
            .tables
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB table catalog",
                source,
            })?;
        catalog.tables.push(Table::parsed(name, table_kind));
        table = Some(catalog.tables.len() - 1);
    }
    if in_database && matches!(kind, Element::Column | Element::ColumnDefinition) && table.is_some()
    {
        add_column(reader, element, table, catalog, columns, limits)?;
    }
    if in_database && kind == Element::Query {
        ensure_capacity(
            catalog.queries.len(),
            limits.max_queries,
            "query declarations",
        )?;
        let name = required_db_attr(reader, element, b"name", limits)?;
        let command = required_db_attr(reader, element, b"command", limits)?;
        let escape_processing = optional_db_attr(reader, element, b"escape-processing", limits)?
            .map(|value| parse_bool(&value))
            .transpose()?;
        catalog
            .queries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB query catalog",
                source,
            })?;
        catalog
            .queries
            .push(Query::parsed(name, command, escape_processing));
    }

    Ok(Frame {
        element: kind,
        in_database,
        table,
    })
}

fn add_column(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    table: Option<usize>,
    catalog: &mut OwnedCatalog,
    columns: &mut usize,
    limits: Limits,
) -> Result<()> {
    ensure_capacity(*columns, limits.max_columns, "column declarations")?;
    let name = required_db_attr(reader, element, b"name", limits)?;
    let index = table.ok_or_else(|| invalid("ODB column has no table owner"))?;
    let target = catalog
        .tables
        .get_mut(index)
        .ok_or_else(|| invalid("ODB column table owner is out of bounds"))?;
    target.try_push_column(Column::parsed(name))?;
    *columns = columns
        .checked_add(1)
        .ok_or_else(|| invalid("ODB column count overflow"))?;
    Ok(())
}

fn classify(parent: Option<Frame>, namespace: NamespaceKind, local: &[u8]) -> Element {
    let office = namespace == NamespaceKind::Office;
    let database = namespace == NamespaceKind::Database;
    match (parent.map(|frame| frame.element), office, database, local) {
        (None, true, _, b"document-content") => Element::Document,
        (Some(Element::Document), true, _, b"body") => Element::Body,
        (Some(Element::Body), true, _, b"database") => Element::Database,
        (Some(Element::Database), _, true, b"data-source") => Element::DataSource,
        (Some(Element::Database), _, true, b"queries") => Element::Queries,
        (Some(Element::Queries | Element::QueryCollection), _, true, b"query-collection") => {
            Element::QueryCollection
        },
        (Some(Element::Queries | Element::QueryCollection), _, true, b"query") => Element::Query,
        (Some(Element::Database), _, true, b"table-representations") => {
            Element::TableRepresentations
        },
        (Some(Element::TableRepresentations), _, true, b"table-representation") => {
            Element::TableRepresentation
        },
        (Some(Element::TableRepresentation | Element::Query), _, true, b"columns") => {
            Element::Columns
        },
        (Some(Element::Columns), _, true, b"column") => Element::Column,
        (Some(Element::Database), _, true, b"schema-definition") => Element::SchemaDefinition,
        (Some(Element::SchemaDefinition), _, true, b"table-definitions") => {
            Element::TableDefinitions
        },
        (Some(Element::TableDefinitions), _, true, b"table-definition") => Element::TableDefinition,
        (Some(Element::TableDefinition), _, true, b"column-definitions") => {
            Element::ColumnDefinitions
        },
        (Some(Element::ColumnDefinitions), _, true, b"column-definition") => {
            Element::ColumnDefinition
        },
        _ => Element::Other,
    }
}

fn is_catalog_node(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Database
        && matches!(
            local,
            b"data-source"
                | b"queries"
                | b"query-collection"
                | b"query"
                | b"table-representations"
                | b"table-representation"
                | b"columns"
                | b"column"
                | b"schema-definition"
                | b"table-definitions"
                | b"table-definition"
                | b"column-definitions"
                | b"column-definition"
        )
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE) {
        NamespaceKind::Office
    } else if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == DATABASE_NAMESPACE)
    {
        NamespaceKind::Database
    } else {
        NamespaceKind::Other
    }
}

fn required_db_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    limits: Limits,
) -> Result<String> {
    optional_db_attr(reader, element, local, limits)?.ok_or_else(|| {
        invalid(&format!(
            "ODB {} is missing db:{}",
            String::from_utf8_lossy(element.local_name().as_ref()),
            String::from_utf8_lossy(local)
        ))
    })
}

fn optional_db_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    limits: Limits,
) -> Result<Option<String>> {
    let mut found = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(&format!("invalid ODB attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE)
            && name.as_ref() == local
        {
            if attribute.value.len() > limits.max_attribute_bytes {
                return Err(invalid("ODB semantic attribute exceeds the byte limit"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| invalid(&format!("invalid ODB attribute value: {error}")))?
                .into_owned();
            if value.len() > limits.max_attribute_bytes {
                return Err(invalid(
                    "decoded ODB semantic attribute exceeds the byte limit",
                ));
            }
            if found.replace(value).is_some() {
                return Err(invalid("duplicate ODB semantic attribute"));
            }
        }
    }
    Ok(found)
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("invalid ODB boolean attribute")),
    }
}

fn ensure_capacity(current: usize, limit: usize, what: &str) -> Result<()> {
    if current >= limit {
        return Err(invalid(&format!(
            "ODB semantic catalog exceeds the {what} limit"
        )));
    }
    Ok(())
}

fn select<'a, T>(
    values: &'a [T],
    name: &str,
    name_of: impl Fn(&T) -> &str,
    kind: &str,
) -> Result<Option<&'a T>> {
    let mut selected = None;
    for value in values {
        if name_of(value) == name && selected.replace(value).is_some() {
            return Err(invalid(&format!("ODB {kind} name '{name}' is ambiguous")));
        }
    }
    Ok(selected)
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}
