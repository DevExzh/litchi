//! Bounded, inert ODB schema and query catalogs.

use super::{
    component::{Component, ComponentKind},
    connection::Connection,
    query::Query,
    table::{
        Column, ColumnSchema, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind,
        ReferentialAction, Relation, Table, TableKind,
    },
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
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";

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
    max_components: usize,
    max_keys: usize,
    max_indices: usize,
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

    /// Sets the maximum combined form and report component declarations.
    #[must_use]
    pub const fn with_max_components(mut self, value: usize) -> Self {
        self.max_components = value;
        self
    }

    /// Sets the maximum key declarations.
    #[must_use]
    pub const fn with_max_keys(mut self, value: usize) -> Self {
        self.max_keys = value;
        self
    }

    /// Sets the maximum index declarations.
    #[must_use]
    pub const fn with_max_indices(mut self, value: usize) -> Self {
        self.max_indices = value;
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
            max_components: 65_536,
            max_keys: 65_536,
            max_indices: 65_536,
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

    /// Returns inert form and report declarations in source order.
    #[must_use]
    pub fn components(&self) -> &[Component] {
        self.owned.components()
    }

    /// Returns foreign-key relations in table/key source order.
    #[must_use]
    pub fn relations(&self) -> &[Relation] {
        self.owned.relations()
    }

    /// Returns the inert connection declaration, if the data source has one.
    #[must_use]
    pub const fn connection(&self) -> Option<&Connection> {
        self.owned.connection()
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
    components: Vec<Component>,
    relations: Vec<Relation>,
    connection: Option<Connection>,
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

    /// Returns inert form and report declarations in source order.
    #[must_use]
    pub fn components(&self) -> &[Component] {
        &self.components
    }

    /// Returns foreign-key relations in table/key source order.
    #[must_use]
    pub fn relations(&self) -> &[Relation] {
        &self.relations
    }

    /// Returns the inert connection declaration, if the data source has one.
    #[must_use]
    pub const fn connection(&self) -> Option<&Connection> {
        self.connection.as_ref()
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Element {
    Document,
    Body,
    Database,
    DataSource,
    ConnectionData,
    DatabaseDescription,
    FileBasedDatabase,
    ServerDatabase,
    ConnectionResource,
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
    Forms,
    Reports,
    ComponentCollection,
    Component,
    Keys,
    Key,
    KeyColumns,
    KeyColumn,
    Indices,
    Index,
    IndexColumns,
    IndexColumn,
    Other,
}

#[derive(Clone, Copy)]
struct Frame {
    element: Element,
    in_database: bool,
    table: Option<usize>,
    component_kind: Option<ComponentKind>,
    key: Option<usize>,
    index: Option<usize>,
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
    let mut components = 0usize;
    let mut keys = 0usize;
    let mut indices = 0usize;

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
                    &mut components,
                    &mut keys,
                    &mut indices,
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
                    &mut components,
                    &mut keys,
                    &mut indices,
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
    catalog.relations = collect_relations(&catalog.tables)?;
    Ok(catalog)
}

fn collect_relations(tables: &[Table]) -> Result<Vec<Relation>> {
    let count = tables
        .iter()
        .map(|table| {
            table
                .keys()
                .iter()
                .filter(|key| key.kind() == KeyKind::Foreign && key.referenced_table().is_some())
                .count()
        })
        .try_fold(0usize, usize::checked_add)
        .ok_or_else(|| invalid("ODB relation count overflow"))?;
    let mut relations = Vec::new();
    relations
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation {
            resource: "ODB relation catalog",
            source,
        })?;
    for table in tables {
        for key in table.keys() {
            if key.kind() == KeyKind::Foreign
                && let Some(relation) = Relation::from_key(table.name(), key)
            {
                relations.push(relation);
            }
        }
    }
    Ok(relations)
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
    components: &mut usize,
    keys: &mut usize,
    indices: &mut usize,
    limits: Limits,
) -> Result<Frame> {
    let local = element.local_name();
    let kind = classify(parent, namespace, local.as_ref());
    let in_database = parent.is_some_and(|frame| frame.in_database) || kind == Element::Database;
    if kind == Element::Other && is_catalog_node(namespace, local.as_ref()) {
        return Err(invalid("ODB schema node has an invalid parent"));
    }
    let mut table = parent.and_then(|frame| frame.table);
    let mut key = parent.and_then(|frame| frame.key);
    let mut index = parent.and_then(|frame| frame.index);
    let component_kind = match kind {
        Element::Forms => Some(ComponentKind::Form),
        Element::Reports => Some(ComponentKind::Report),
        Element::ComponentCollection | Element::Component => {
            parent.and_then(|frame| frame.component_kind)
        },
        Element::Document
        | Element::Body
        | Element::Database
        | Element::DataSource
        | Element::ConnectionData
        | Element::DatabaseDescription
        | Element::FileBasedDatabase
        | Element::ServerDatabase
        | Element::ConnectionResource
        | Element::Queries
        | Element::QueryCollection
        | Element::TableRepresentation
        | Element::TableRepresentations
        | Element::TableDefinition
        | Element::TableDefinitions
        | Element::SchemaDefinition
        | Element::Columns
        | Element::ColumnDefinitions
        | Element::Column
        | Element::ColumnDefinition
        | Element::Query
        | Element::Keys
        | Element::Key
        | Element::KeyColumns
        | Element::KeyColumn
        | Element::Indices
        | Element::Index
        | Element::IndexColumns
        | Element::IndexColumn
        | Element::Other => None,
    };

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
            Element::FileBasedDatabase | Element::ServerDatabase | Element::ConnectionResource
        )
    {
        let connection = match kind {
            Element::FileBasedDatabase => Connection::file(required_attr(
                reader,
                element,
                XLINK_NAMESPACE,
                b"href",
                limits,
            )?),
            Element::ServerDatabase => Connection::server(
                required_db_attr(reader, element, b"hostname", limits)?,
                required_db_attr(reader, element, b"database-name", limits)?,
            ),
            Element::ConnectionResource => Connection::resource(required_attr(
                reader,
                element,
                XLINK_NAMESPACE,
                b"href",
                limits,
            )?),
            Element::Document
            | Element::Body
            | Element::Database
            | Element::DataSource
            | Element::ConnectionData
            | Element::DatabaseDescription
            | Element::Queries
            | Element::QueryCollection
            | Element::TableRepresentation
            | Element::TableRepresentations
            | Element::TableDefinition
            | Element::TableDefinitions
            | Element::SchemaDefinition
            | Element::Columns
            | Element::ColumnDefinitions
            | Element::Column
            | Element::ColumnDefinition
            | Element::Query
            | Element::Forms
            | Element::Reports
            | Element::ComponentCollection
            | Element::Component
            | Element::Keys
            | Element::Key
            | Element::KeyColumns
            | Element::KeyColumn
            | Element::Indices
            | Element::Index
            | Element::IndexColumns
            | Element::IndexColumn
            | Element::Other => {
                return Err(invalid("ODB connection classification is inconsistent"));
            },
        };
        if catalog.connection.replace(connection).is_some() {
            return Err(invalid(
                "ODB data source has more than one connection declaration",
            ));
        }
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
        key = None;
        index = None;
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
    if in_database && kind == Element::Component {
        ensure_capacity(*components, limits.max_components, "component declarations")?;
        let owner =
            component_kind.ok_or_else(|| invalid("ODB component has no form/report owner"))?;
        let as_template = optional_db_attr(reader, element, b"as-template", limits)?
            .map(|value| parse_bool(&value))
            .transpose()?;
        let component = Component::parsed(
            owner,
            optional_db_attr(reader, element, b"name", limits)?,
            optional_db_attr(reader, element, b"title", limits)?,
            optional_db_attr(reader, element, b"description", limits)?,
            optional_attr(reader, element, XLINK_NAMESPACE, b"href", limits)?,
            as_template,
        );
        catalog
            .components
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB form/report component catalog",
                source,
            })?;
        catalog.components.push(component);
        *components = components
            .checked_add(1)
            .ok_or_else(|| invalid("ODB component count overflow"))?;
    }
    if in_database && kind == Element::Key {
        ensure_capacity(*keys, limits.max_keys, "key declarations")?;
        let table_index = table.ok_or_else(|| invalid("ODB key has no table owner"))?;
        let key_kind = parse_key_kind(&required_db_attr(reader, element, b"type", limits)?)?;
        let update_rule = optional_db_attr(reader, element, b"update-rule", limits)?
            .map(|value| parse_referential_action(&value))
            .transpose()?;
        let delete_rule = optional_db_attr(reader, element, b"delete-rule", limits)?
            .map(|value| parse_referential_action(&value))
            .transpose()?;
        let target = catalog
            .tables
            .get_mut(table_index)
            .ok_or_else(|| invalid("ODB key table owner is out of bounds"))?;
        target.try_push_key(Key::parsed(
            optional_db_attr(reader, element, b"name", limits)?,
            key_kind,
            optional_db_attr(reader, element, b"referenced-table-name", limits)?,
            update_rule,
            delete_rule,
        ))?;
        key = target.keys().len().checked_sub(1);
        *keys = keys
            .checked_add(1)
            .ok_or_else(|| invalid("ODB key count overflow"))?;
    }
    if in_database && kind == Element::KeyColumn {
        let table_index = table.ok_or_else(|| invalid("ODB key column has no table owner"))?;
        let key_index = key.ok_or_else(|| invalid("ODB key column has no key owner"))?;
        let target = catalog
            .tables
            .get_mut(table_index)
            .and_then(|table_value| table_value.keys_mut().get_mut(key_index))
            .ok_or_else(|| invalid("ODB key column owner is out of bounds"))?;
        target.try_push_column(KeyColumn::parsed(
            optional_db_attr(reader, element, b"name", limits)?,
            optional_db_attr(reader, element, b"related-column-name", limits)?,
        ))?;
    }
    if in_database && kind == Element::Index {
        ensure_capacity(*indices, limits.max_indices, "index declarations")?;
        let table_index = table.ok_or_else(|| invalid("ODB index has no table owner"))?;
        let unique = optional_db_attr(reader, element, b"is-unique", limits)?
            .map(|value| parse_bool(&value))
            .transpose()?;
        let clustered = optional_db_attr(reader, element, b"is-clustered", limits)?
            .map(|value| parse_bool(&value))
            .transpose()?;
        let target = catalog
            .tables
            .get_mut(table_index)
            .ok_or_else(|| invalid("ODB index table owner is out of bounds"))?;
        target.try_push_index(Index::parsed(
            required_db_attr(reader, element, b"name", limits)?,
            unique,
            clustered,
        ))?;
        index = target.indices().len().checked_sub(1);
        *indices = indices
            .checked_add(1)
            .ok_or_else(|| invalid("ODB index count overflow"))?;
    }
    if in_database && kind == Element::IndexColumn {
        let table_index = table.ok_or_else(|| invalid("ODB index column has no table owner"))?;
        let index_value = index.ok_or_else(|| invalid("ODB index column has no index owner"))?;
        let ascending = optional_db_attr(reader, element, b"is-ascending", limits)?
            .map(|value| parse_bool(&value))
            .transpose()?;
        let target = catalog
            .tables
            .get_mut(table_index)
            .and_then(|table_value| table_value.indices_mut().get_mut(index_value))
            .ok_or_else(|| invalid("ODB index column owner is out of bounds"))?;
        target.try_push_column(IndexColumn::parsed(
            required_db_attr(reader, element, b"name", limits)?,
            ascending,
        ))?;
    }

    Ok(Frame {
        element: kind,
        in_database,
        table,
        component_kind,
        key,
        index,
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
    let data_type_value = optional_db_attr(reader, element, b"data-type", limits)?;
    let data_type = data_type_value
        .as_deref()
        .map(validate_data_type)
        .transpose()?;
    let precision = optional_db_attr(reader, element, b"precision", limits)?
        .map(|value| parse_positive_integer(&value, "precision"))
        .transpose()?;
    let scale = optional_db_attr(reader, element, b"scale", limits)?
        .map(|value| parse_positive_integer(&value, "scale"))
        .transpose()?;
    let nullable = optional_db_attr(reader, element, b"is-nullable", limits)?
        .map(|value| parse_nullability(&value))
        .transpose()?;
    let empty_allowed = optional_db_attr(reader, element, b"is-empty-allowed", limits)?
        .map(|value| parse_bool(&value))
        .transpose()?;
    let autoincrement = optional_db_attr(reader, element, b"is-autoincrement", limits)?
        .map(|value| parse_bool(&value))
        .transpose()?;
    target.try_push_column(Column::parsed(
        name,
        ColumnSchema {
            data_type,
            type_name: optional_db_attr(reader, element, b"type-name", limits)?,
            precision,
            scale,
            nullable,
            empty_allowed,
            autoincrement,
        },
    ))?;
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
        (Some(Element::DataSource), _, true, b"connection-data") => Element::ConnectionData,
        (Some(Element::ConnectionData), _, true, b"database-description") => {
            Element::DatabaseDescription
        },
        (Some(Element::DatabaseDescription), _, true, b"file-based-database") => {
            Element::FileBasedDatabase
        },
        (Some(Element::DatabaseDescription), _, true, b"server-database") => {
            Element::ServerDatabase
        },
        (Some(Element::ConnectionData), _, true, b"connection-resource") => {
            Element::ConnectionResource
        },
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
        (Some(Element::TableDefinition), _, true, b"keys") => Element::Keys,
        (Some(Element::Keys), _, true, b"key") => Element::Key,
        (Some(Element::Key), _, true, b"key-columns") => Element::KeyColumns,
        (Some(Element::KeyColumns), _, true, b"key-column") => Element::KeyColumn,
        (Some(Element::TableDefinition), _, true, b"indices") => Element::Indices,
        (Some(Element::Indices), _, true, b"index") => Element::Index,
        (Some(Element::Index), _, true, b"index-columns") => Element::IndexColumns,
        (Some(Element::IndexColumns), _, true, b"index-column") => Element::IndexColumn,
        (Some(Element::Database), _, true, b"forms") => Element::Forms,
        (Some(Element::Database), _, true, b"reports") => Element::Reports,
        (
            Some(Element::Forms | Element::Reports | Element::ComponentCollection),
            _,
            true,
            b"component-collection",
        ) => Element::ComponentCollection,
        (
            Some(Element::Forms | Element::Reports | Element::ComponentCollection),
            _,
            true,
            b"component",
        ) => Element::Component,
        _ => Element::Other,
    }
}

fn is_catalog_node(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Database
        && matches!(
            local,
            b"data-source"
                | b"connection-data"
                | b"database-description"
                | b"file-based-database"
                | b"server-database"
                | b"connection-resource"
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
                | b"forms"
                | b"reports"
                | b"component-collection"
                | b"component"
                | b"keys"
                | b"key"
                | b"key-columns"
                | b"key-column"
                | b"indices"
                | b"index"
                | b"index-columns"
                | b"index-column"
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

fn required_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
    limits: Limits,
) -> Result<String> {
    optional_attr(reader, element, namespace, local, limits)?.ok_or_else(|| {
        invalid(&format!(
            "ODB {} is missing required attribute {}",
            String::from_utf8_lossy(element.local_name().as_ref()),
            String::from_utf8_lossy(local)
        ))
    })
}

fn optional_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local: &[u8],
    limits: Limits,
) -> Result<Option<String>> {
    let mut found = None;
    for raw_attribute in element.attributes() {
        let attribute =
            raw_attribute.map_err(|error| invalid(&format!("invalid ODB attribute: {error}")))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == expected_namespace)
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

fn parse_key_kind(value: &str) -> Result<KeyKind> {
    match value {
        "primary" => Ok(KeyKind::Primary),
        "unique" => Ok(KeyKind::Unique),
        "foreign" => Ok(KeyKind::Foreign),
        _ => Err(invalid("invalid ODB key type")),
    }
}

fn parse_referential_action(value: &str) -> Result<ReferentialAction> {
    match value {
        "cascade" => Ok(ReferentialAction::Cascade),
        "restrict" => Ok(ReferentialAction::Restrict),
        "set-null" => Ok(ReferentialAction::SetNull),
        "no-action" => Ok(ReferentialAction::NoAction),
        "set-default" => Ok(ReferentialAction::SetDefault),
        _ => Err(invalid("invalid ODB referential action")),
    }
}

fn parse_nullability(value: &str) -> Result<bool> {
    match value {
        "nullable" => Ok(true),
        "no-nulls" => Ok(false),
        _ => Err(invalid("invalid ODB column nullability")),
    }
}

fn parse_positive_integer(value: &str, kind: &str) -> Result<u64> {
    let parsed = value
        .parse::<u64>()
        .map_err(|_error| invalid(&format!("invalid ODB column {kind}")))?;
    if parsed == 0 {
        return Err(invalid(&format!("ODB column {kind} must be positive")));
    }
    Ok(parsed)
}

fn validate_data_type(value: &str) -> Result<DataType> {
    match value {
        "bit" => Ok(DataType::Bit),
        "boolean" => Ok(DataType::Boolean),
        "tinyint" => Ok(DataType::TinyInt),
        "smallint" => Ok(DataType::SmallInt),
        "integer" => Ok(DataType::Integer),
        "bigint" => Ok(DataType::BigInt),
        "float" => Ok(DataType::Float),
        "real" => Ok(DataType::Real),
        "double" => Ok(DataType::Double),
        "numeric" => Ok(DataType::Numeric),
        "decimal" => Ok(DataType::Decimal),
        "char" => Ok(DataType::Char),
        "varchar" => Ok(DataType::VarChar),
        "longvarchar" => Ok(DataType::LongVarChar),
        "date" => Ok(DataType::Date),
        "time" => Ok(DataType::Time),
        "timestmp" => Ok(DataType::Timestamp),
        "binary" => Ok(DataType::Binary),
        "varbinary" => Ok(DataType::VarBinary),
        "longvarbinary" => Ok(DataType::LongVarBinary),
        "sqlnull" => Ok(DataType::SqlNull),
        "other" => Ok(DataType::Other),
        "object" => Ok(DataType::Object),
        "distinct" => Ok(DataType::Distinct),
        "struct" => Ok(DataType::Struct),
        "array" => Ok(DataType::Array),
        "blob" => Ok(DataType::Blob),
        "clob" => Ok(DataType::Clob),
        "ref" => Ok(DataType::Ref),
        _ => Err(invalid("invalid ODB column data type")),
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
