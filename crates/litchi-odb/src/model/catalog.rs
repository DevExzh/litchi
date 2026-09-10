//! Bounded, inert ODB schema and query catalogs.

use super::{
    component::{Component, ComponentKind},
    connection::{Connection, FileDatabaseTarget, ServerDatabaseTarget},
    query::{Query, QueryUpdateTarget},
    settings::{self, DatabaseSettings},
    table::{
        Column, ColumnSchema, DataType, Index, IndexColumn, Key, KeyColumn, KeyKind,
        ReferentialAction, Relation, RelationResolution, Table, TableKind,
    },
};
use litchi_core::{Error, Result};
use quick_xml::{
    events::{BytesStart, Event, attributes::Attribute},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use std::{
    fmt,
    sync::{Arc, OnceLock},
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
    max_settings: usize,
    max_setting_values: usize,
    max_table_settings: usize,
    max_filter_patterns: usize,
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

    /// Sets the maximum number of data-source setting declarations.
    #[must_use]
    pub const fn with_max_settings(mut self, value: usize) -> Self {
        self.max_settings = value;
        self
    }

    /// Sets the maximum number of data-source setting values.
    #[must_use]
    pub const fn with_max_setting_values(mut self, value: usize) -> Self {
        self.max_setting_values = value;
        self
    }

    /// Sets the maximum number of driver table-setting declarations.
    #[must_use]
    pub const fn with_max_table_settings(mut self, value: usize) -> Self {
        self.max_table_settings = value;
        self
    }

    /// Sets the maximum combined table-filter patterns and table types.
    #[must_use]
    pub const fn with_max_filter_patterns(mut self, value: usize) -> Self {
        self.max_filter_patterns = value;
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
            max_settings: 65_536,
            max_setting_values: 65_536,
            max_table_settings: 65_536,
            max_filter_patterns: 65_536,
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
    settings: LazySettings<'source>,
}

/// A source-backed settings projection that materializes only after an
/// explicit settings read. Failed parses are deliberately not cached, so a
/// caller never observes a stale error after the source-bound view changes.
#[derive(Clone)]
struct LazySettings<'source> {
    source: &'source str,
    limits: Limits,
    value: Arc<OnceLock<DatabaseSettings>>,
}

impl<'source> LazySettings<'source> {
    fn with_cache(
        source: &'source str,
        limits: Limits,
        value: Arc<OnceLock<DatabaseSettings>>,
    ) -> Self {
        Self {
            source,
            limits,
            value,
        }
    }

    fn get(&self) -> Result<&DatabaseSettings> {
        if let Some(value) = self.value.get() {
            return Ok(value);
        }
        let value = settings::parse(
            self.source,
            self.limits.max_attribute_bytes,
            self.limits.max_settings,
            self.limits.max_setting_values,
            self.limits.max_table_settings,
            self.limits.max_filter_patterns,
        )?;
        let _ = self.value.set(value);
        self.value
            .get()
            .ok_or_else(|| invalid("ODB settings projection was not materialized"))
    }
}

impl fmt::Debug for LazySettings<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("LazySettings")
            .field("source_bytes", &self.source.len())
            .field("limits", &self.limits)
            .field("materialized", &self.value.get().is_some())
            .finish()
    }
}

impl PartialEq for LazySettings<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source && self.limits == other.limits
    }
}

impl Eq for LazySettings<'_> {}

impl<'source> Catalog<'source> {
    pub(crate) fn parse(source: &'source str, limits: Limits) -> Result<Self> {
        Self::parse_with_settings(source, limits, Arc::new(OnceLock::new()))
    }

    pub(crate) fn parse_with_settings(
        source: &'source str,
        limits: Limits,
        settings_cache: Arc<OnceLock<DatabaseSettings>>,
    ) -> Result<Self> {
        Ok(Self {
            source,
            owned: parse(source, limits)?,
            settings: LazySettings::with_cache(source, limits, settings_cache),
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

    /// Iterates foreign-key relations owned by `table`.
    pub fn outgoing_relations<'catalog>(
        &'catalog self,
        table: &'catalog str,
    ) -> impl Iterator<Item = &'catalog Relation> + 'catalog {
        self.relations()
            .iter()
            .filter(move |relation| relation.table() == table)
    }

    /// Iterates foreign-key relations targeting `table`.
    pub fn incoming_relations<'catalog>(
        &'catalog self,
        table: &'catalog str,
    ) -> impl Iterator<Item = &'catalog Relation> + 'catalog {
        self.relations()
            .iter()
            .filter(move |relation| relation.referenced_table() == table)
    }

    /// Resolves one relation against this inert local catalog.
    #[must_use]
    pub fn relation_resolution(&self, relation: &Relation) -> RelationResolution {
        let Some(owner) = self
            .tables()
            .iter()
            .find(|table| table.name() == relation.table())
        else {
            return RelationResolution::MissingOwnerTable;
        };
        let Some(target) = self
            .tables()
            .iter()
            .find(|table| table.name() == relation.referenced_table())
        else {
            return RelationResolution::MissingReferencedTable;
        };
        for mapping in relation.columns() {
            let Some(local) = mapping.name() else {
                return RelationResolution::IncompleteLocalColumn;
            };
            if owner.column(local).is_none() {
                return RelationResolution::MissingLocalColumn;
            }
            if let Some(related) = mapping.related_column()
                && target.column(related).is_none()
            {
                return RelationResolution::MissingReferencedColumn;
            }
        }
        RelationResolution::Resolved
    }

    /// Returns the inert connection declaration, if the data source has one.
    #[must_use]
    pub const fn connection(&self) -> Option<&Connection> {
        self.owned.connection()
    }

    /// Returns bounded inert login, driver, application, and filter settings.
    ///
    /// The optional projection is parsed on the first call and successful
    /// reads are reused by this source-bound view. A malformed projection is
    /// returned as a typed error and is not cached.
    pub fn settings(&self) -> Result<&DatabaseSettings> {
        self.settings.get()
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
    FilterStatement,
    OrderStatement,
    UpdateTable,
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
    query: Option<usize>,
    component_kind: Option<ComponentKind>,
    key: Option<usize>,
    index: Option<usize>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Database,
    Xlink,
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
    let mut strict_targets = false;

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
        match raw_event {
            Event::Start(element) => {
                if stack
                    .last()
                    .is_some_and(|frame| is_connection_target(frame.element))
                {
                    return Err(invalid("ODB connection target must have empty content"));
                }
                let event_strict_targets = if stack.is_empty()
                    && namespace == NamespaceKind::Office
                    && element.local_name().as_ref() == b"document-content"
                {
                    odf14_targets(&reader, &element, limits)?
                } else {
                    strict_targets
                };
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
                    event_strict_targets,
                    limits,
                )?;
                if stack.is_empty() {
                    if root_seen || root_closed || frame.element != Element::Document {
                        return Err(invalid(
                            "ODB semantic catalog has no office:document-content root",
                        ));
                    }
                    root_seen = true;
                    strict_targets = event_strict_targets;
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
                if stack
                    .last()
                    .is_some_and(|frame| is_connection_target(frame.element))
                {
                    return Err(invalid("ODB connection target must have empty content"));
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
                    strict_targets,
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
            Event::Text(_) | Event::CData(_)
                if stack
                    .last()
                    .is_some_and(|frame| is_connection_target(frame.element)) =>
            {
                return Err(invalid("ODB connection target must have empty content"));
            },
            Event::GeneralRef(_)
                if stack
                    .last()
                    .is_some_and(|frame| is_connection_target(frame.element)) =>
            {
                return Err(invalid("ODB connection target must have empty content"));
            },
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
    strict_targets: bool,
    limits: Limits,
) -> Result<Frame> {
    let local = element.local_name();
    validate_reserved_element(parent.map(|frame| frame.element), namespace, local.as_ref())?;
    validate_reserved_attributes(reader, element, local.as_ref())?;
    let kind = classify(parent, namespace, local.as_ref());
    let in_database = parent.is_some_and(|frame| frame.in_database) || kind == Element::Database;
    if kind == Element::Other && is_catalog_node(namespace, local.as_ref()) {
        return Err(invalid("ODB schema node has an invalid parent"));
    }
    let mut table = parent.and_then(|frame| frame.table);
    let mut query = parent.and_then(|frame| frame.query);
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
        | Element::FilterStatement
        | Element::OrderStatement
        | Element::UpdateTable
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
            Element::FileBasedDatabase => {
                parse_file_target(reader, element, strict_targets, limits)?
            },
            Element::ServerDatabase => {
                parse_server_target(reader, element, strict_targets, limits)?
            },
            Element::ConnectionResource => parse_resource_target(reader, element, limits)?,
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
            | Element::FilterStatement
            | Element::OrderStatement
            | Element::UpdateTable
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
        query = None;
        key = None;
        index = None;
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
        query = catalog.queries.len().checked_sub(1);
        table = None;
    }
    if in_database
        && matches!(kind, Element::Column | Element::ColumnDefinition)
        && (table.is_some() || query.is_some())
    {
        add_column(reader, element, table, query, catalog, columns, limits)?;
    }
    if in_database && matches!(kind, Element::FilterStatement | Element::OrderStatement) {
        let command = required_db_attr(reader, element, b"command", limits)?;
        match (table, query, kind) {
            (Some(index), None, Element::FilterStatement) => catalog
                .tables
                .get_mut(index)
                .ok_or_else(|| invalid("ODB filter table owner is out of bounds"))?
                .set_filter_statement(command)?,
            (Some(index), None, Element::OrderStatement) => catalog
                .tables
                .get_mut(index)
                .ok_or_else(|| invalid("ODB order table owner is out of bounds"))?
                .set_order_statement(command)?,
            (None, Some(index), Element::FilterStatement) => catalog
                .queries
                .get_mut(index)
                .ok_or_else(|| invalid("ODB filter query owner is out of bounds"))?
                .set_filter_statement(command)?,
            (None, Some(index), Element::OrderStatement) => catalog
                .queries
                .get_mut(index)
                .ok_or_else(|| invalid("ODB order query owner is out of bounds"))?
                .set_order_statement(command)?,
            _ => return Err(invalid("ODB statement owner is ambiguous")),
        }
    }
    if in_database && kind == Element::UpdateTable {
        let index = query.ok_or_else(|| invalid("ODB update-table has no query owner"))?;
        catalog
            .queries
            .get_mut(index)
            .ok_or_else(|| invalid("ODB update-table query owner is out of bounds"))?
            .set_update_target(QueryUpdateTarget::parsed(
                required_db_attr(reader, element, b"name", limits)?,
                optional_db_attr(reader, element, b"schema-name", limits)?,
                optional_db_attr(reader, element, b"catalog-name", limits)?,
            ))?;
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
        query,
        component_kind,
        key,
        index,
    })
}

fn add_column(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    table: Option<usize>,
    query: Option<usize>,
    catalog: &mut OwnedCatalog,
    columns: &mut usize,
    limits: Limits,
) -> Result<()> {
    ensure_capacity(*columns, limits.max_columns, "column declarations")?;
    let name = required_db_attr(reader, element, b"name", limits)?;
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
    let column = Column::parsed(
        name,
        ColumnSchema {
            data_type,
            type_name: optional_db_attr(reader, element, b"type-name", limits)?,
            precision,
            scale,
            nullable,
            empty_allowed,
            autoincrement,
            default_value: optional_db_attr(reader, element, b"default-value", limits)?,
        },
    );
    match (table, query) {
        (Some(index), None) => catalog
            .tables
            .get_mut(index)
            .ok_or_else(|| invalid("ODB column table owner is out of bounds"))?
            .try_push_column(column)?,
        (None, Some(index)) => catalog
            .queries
            .get_mut(index)
            .ok_or_else(|| invalid("ODB column query owner is out of bounds"))?
            .try_push_column(column)?,
        (None, None) | (Some(_), Some(_)) => {
            return Err(invalid("ODB column owner is ambiguous"));
        },
    }
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
        (Some(Element::TableRepresentation | Element::Query), _, true, b"filter-statement") => {
            Element::FilterStatement
        },
        (Some(Element::TableRepresentation | Element::Query), _, true, b"order-statement") => {
            Element::OrderStatement
        },
        (Some(Element::Query), _, true, b"update-table") => Element::UpdateTable,
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
                | b"filter-statement"
                | b"order-statement"
                | b"update-table"
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

fn is_connection_target(element: Element) -> bool {
    matches!(
        element,
        Element::FileBasedDatabase | Element::ServerDatabase | Element::ConnectionResource
    )
}

fn odf14_targets(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
) -> Result<bool> {
    let Some(version) = optional_attr(reader, element, OFFICE_NAMESPACE, b"version", limits)?
    else {
        return Ok(false);
    };
    let version = collapse_xml_whitespace(&version)?;
    let mut parts = version.split('.');
    let major = parts
        .next()
        .ok_or_else(|| invalid("ODB office:version is empty"))?
        .parse::<u16>()
        .map_err(|_| invalid("invalid ODB office:version"))?;
    let minor = parts
        .next()
        .ok_or_else(|| invalid("invalid ODB office:version"))?
        .parse::<u16>()
        .map_err(|_| invalid("invalid ODB office:version"))?;
    if parts.next().is_some() {
        return Err(invalid("invalid ODB office:version"));
    }
    Ok((major, minor) >= (1, 4))
}

fn parse_file_target(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    strict_targets: bool,
    limits: Limits,
) -> Result<Connection> {
    validate_connection_target_attributes(reader, element, b"file-based-database")?;
    let href = required_attr(reader, element, XLINK_NAMESPACE, b"href", limits)?;
    // `xlink:type` is a fixed schema token.  Keep its tiny lexical value
    // readable even when a caller sets the semantic string limit to zero to
    // exercise the separately lazy settings projection.
    let link_type = optional_attr(
        reader,
        element,
        XLINK_NAMESPACE,
        b"type",
        connection_link_limits(limits),
    )?;
    let media_type = optional_db_attr(reader, element, b"media-type", limits)?;
    let extension = optional_db_attr(reader, element, b"extension", limits)?;
    if let Some(value) = link_type.as_deref()
        && collapse_xml_whitespace(value)? != "simple"
    {
        return Err(invalid("ODB file-based-database xlink:type must be simple"));
    }
    if strict_targets && (link_type.is_none() || media_type.is_none()) {
        return Err(invalid(
            "ODB file-based-database requires xlink:type and db:media-type",
        ));
    }
    let (Some(_link_type), Some(media_type)) = (link_type, media_type) else {
        return Ok(Connection::file(href));
    };
    Ok(Connection::FileTarget(
        FileDatabaseTarget::new(href, media_type).with_extension(extension),
    ))
}

fn parse_server_target(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    strict_targets: bool,
    limits: Limits,
) -> Result<Connection> {
    validate_connection_target_attributes(reader, element, b"server-database")?;
    let database_type = optional_db_attr(reader, element, b"type", limits)?;
    if database_type.is_none() {
        if strict_targets {
            return Err(invalid("ODB server-database is missing db:type"));
        }
        return Ok(Connection::server(
            required_db_attr(reader, element, b"hostname", limits)?,
            required_db_attr(reader, element, b"database-name", limits)?,
        ));
    }
    let database_type =
        database_type.ok_or_else(|| invalid("ODB server-database is missing db:type"))?;
    let database_type = collapse_xml_whitespace(&database_type)?;
    let database_type_namespace =
        parse_namespaced_token(reader, &database_type, "server-database type")?;
    let hostname = optional_db_attr(reader, element, b"hostname", limits)?;
    let port = optional_db_attr(reader, element, b"port", limits)?
        .map(|value| parse_positive_u64(&value, "server-database port"))
        .transpose()?;
    let local_socket = optional_db_attr(reader, element, b"local-socket", limits)?;
    if hostname.is_some() && local_socket.is_some() {
        return Err(invalid(
            "ODB server-database cannot declare both hostname and local-socket",
        ));
    }
    if port.is_some() && hostname.is_none() {
        return Err(invalid("ODB server-database port requires a hostname"));
    }
    let database_name = optional_db_attr(reader, element, b"database-name", limits)?;
    let mut target = ServerDatabaseTarget::new(database_type);
    if let Some(namespace) = database_type_namespace {
        target = target.with_database_type_namespace(namespace);
    }
    if let Some(hostname) = hostname {
        target = target.with_host(hostname, port);
    } else if let Some(local_socket) = local_socket {
        target = target.with_local_socket(local_socket);
    }
    Ok(Connection::ServerTarget(
        target.with_database_name(database_name),
    ))
}

fn parse_resource_target(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    limits: Limits,
) -> Result<Connection> {
    validate_connection_target_attributes(reader, element, b"connection-resource")?;
    let link_limits = connection_link_limits(limits);
    let link_type = required_attr(reader, element, XLINK_NAMESPACE, b"type", link_limits)?;
    if collapse_xml_whitespace(&link_type)? != "simple" {
        return Err(invalid("ODB connection-resource xlink:type must be simple"));
    }
    let href = required_attr(reader, element, XLINK_NAMESPACE, b"href", limits)?;
    if let Some(show) = optional_attr(reader, element, XLINK_NAMESPACE, b"show", link_limits)?
        && collapse_xml_whitespace(&show)? != "none"
    {
        return Err(invalid("ODB connection-resource xlink:show must be none"));
    }
    if let Some(actuate) = optional_attr(reader, element, XLINK_NAMESPACE, b"actuate", link_limits)?
        && collapse_xml_whitespace(&actuate)? != "onRequest"
    {
        return Err(invalid(
            "ODB connection-resource xlink:actuate must be onRequest",
        ));
    }
    Ok(Connection::resource(href))
}

fn connection_link_limits(mut limits: Limits) -> Limits {
    limits.max_attribute_bytes = limits.max_attribute_bytes.max(64);
    limits
}

fn validate_connection_target_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    target: &[u8],
) -> Result<()> {
    for raw in element.attributes() {
        let attribute = raw.map_err(|error| invalid(&format!("invalid ODB attribute: {error}")))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        let allowed = match namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => match target {
                b"connection-resource" => {
                    matches!(name.as_ref(), b"type" | b"href" | b"show" | b"actuate")
                },
                b"file-based-database" => matches!(name.as_ref(), b"type" | b"href"),
                b"server-database" => false,
                _ => false,
            },
            ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE => match target {
                b"connection-resource" => false,
                b"file-based-database" => matches!(name.as_ref(), b"media-type" | b"extension"),
                b"server-database" => matches!(
                    name.as_ref(),
                    b"type" | b"hostname" | b"port" | b"local-socket" | b"database-name"
                ),
                _ => false,
            },
            ResolveResult::Bound(Namespace(uri))
                if uri == OFFICE_NAMESPACE || uri == b"http://www.w3.org/2000/xmlns/" =>
            {
                false
            },
            ResolveResult::Unknown(_) | ResolveResult::Unbound if reserved_prefix(raw_name) => {
                false
            },
            ResolveResult::Bound(_) | ResolveResult::Unknown(_) | ResolveResult::Unbound => true,
        };
        if !allowed {
            return Err(invalid(
                "ODB connection target contains an unknown reserved attribute",
            ));
        }
    }
    Ok(())
}

fn validate_reserved_element(
    parent: Option<Element>,
    namespace: NamespaceKind,
    local: &[u8],
) -> Result<()> {
    match namespace {
        NamespaceKind::Office => {
            let allowed = match parent {
                None => local == b"document-content",
                Some(Element::Document) => matches!(
                    local,
                    b"automatic-styles" | b"body" | b"font-face-decls" | b"scripts"
                ),
                Some(Element::Body) => local == b"database",
                _ => false,
            };
            if !allowed {
                return Err(invalid(
                    "ODB semantic catalog contains an invalid office element",
                ));
            }
        },
        NamespaceKind::Database if !is_known_database_element(local) => {
            return Err(invalid(
                "ODB semantic catalog contains an unknown database element",
            ));
        },
        NamespaceKind::Xlink => {
            return Err(invalid(
                "ODB semantic catalog contains an invalid xlink element",
            ));
        },
        NamespaceKind::Database | NamespaceKind::Other => {},
    }
    Ok(())
}

fn is_known_database_element(local: &[u8]) -> bool {
    matches!(
        local,
        b"application-connection-settings"
            | b"auto-increment"
            | b"character-set"
            | b"column"
            | b"column-definition"
            | b"column-definitions"
            | b"columns"
            | b"component"
            | b"component-collection"
            | b"connection-data"
            | b"connection-resource"
            | b"data-source"
            | b"data-source-setting"
            | b"data-source-setting-value"
            | b"data-source-settings"
            | b"database-description"
            | b"delimiter"
            | b"driver-settings"
            | b"file-based-database"
            | b"font-charset"
            | b"filter-statement"
            | b"forms"
            | b"index"
            | b"index-column"
            | b"index-columns"
            | b"indices"
            | b"key"
            | b"key-column"
            | b"key-columns"
            | b"keys"
            | b"login"
            | b"order-statement"
            | b"queries"
            | b"query"
            | b"query-collection"
            | b"reports"
            | b"schema-definition"
            | b"server-database"
            | b"table-definition"
            | b"table-definitions"
            | b"table-exclude-filter"
            | b"table-filter"
            | b"table-filter-pattern"
            | b"table-include-filter"
            | b"table-representation"
            | b"table-representations"
            | b"table-setting"
            | b"table-settings"
            | b"table-type"
            | b"table-type-filter"
            | b"update-table"
    )
}

fn validate_reserved_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<()> {
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| invalid(&format!("invalid ODB semantic attribute: {error}")))?;
        let raw_name = attribute.key.as_ref();
        if raw_name == b"xmlns" || raw_name.starts_with(b"xmlns:") {
            continue;
        }
        let (attribute_namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        let allowed = match attribute_namespace {
            ResolveResult::Bound(Namespace(uri)) if uri == DATABASE_NAMESPACE => {
                known_database_attribute(name.as_ref())
            },
            ResolveResult::Bound(Namespace(uri)) if uri == XLINK_NAMESPACE => match local {
                b"component" => matches!(
                    name.as_ref(),
                    b"actuate" | b"href" | b"show" | b"title" | b"type"
                ),
                b"connection-resource" => {
                    matches!(name.as_ref(), b"actuate" | b"href" | b"show" | b"type")
                },
                b"file-based-database" => matches!(name.as_ref(), b"href" | b"type"),
                _ => false,
            },
            ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE => {
                local == b"document-content" && name.as_ref() == b"version"
            },
            ResolveResult::Unknown(_) | ResolveResult::Unbound if reserved_prefix(raw_name) => {
                false
            },
            ResolveResult::Bound(_) | ResolveResult::Unknown(_) | ResolveResult::Unbound => true,
        };
        if !allowed {
            return Err(invalid(
                "ODB semantic catalog contains an unknown reserved attribute",
            ));
        }
    }
    Ok(())
}

fn known_database_attribute(name: &[u8]) -> bool {
    matches!(
        name,
        b"additional-column-statement"
            | b"append-table-alias-name"
            | b"apply-command"
            | b"as-template"
            | b"base-dn"
            | b"boolean-comparison-mode"
            | b"catalog-name"
            | b"command"
            | b"data-source-setting-is-list"
            | b"data-source-setting-name"
            | b"data-source-setting-type"
            | b"data-type"
            | b"database-name"
            | b"decimal"
            | b"default-cell-style-name"
            | b"default-value"
            | b"default-row-style-name"
            | b"delete-rule"
            | b"description"
            | b"enable-sql92-check"
            | b"encoding"
            | b"escape-processing"
            | b"extension"
            | b"field"
            | b"hostname"
            | b"ignore-driver-privileges"
            | b"is-ascending"
            | b"is-autoincrement"
            | b"is-clustered"
            | b"is-empty-allowed"
            | b"is-first-row-header-line"
            | b"is-nullable"
            | b"is-password-required"
            | b"is-table-name-length-limited"
            | b"is-unique"
            | b"local-socket"
            | b"login-timeout"
            | b"max-row-count"
            | b"media-type"
            | b"name"
            | b"parameter-name-substitution"
            | b"port"
            | b"precision"
            | b"referenced-table-name"
            | b"related-column-name"
            | b"row-retrieving-statement"
            | b"scale"
            | b"schema-name"
            | b"show-deleted"
            | b"string"
            | b"style-name"
            | b"suppress-version-columns"
            | b"system-driver-settings"
            | b"thousand"
            | b"title"
            | b"type"
            | b"type-name"
            | b"update-rule"
            | b"use-catalog"
            | b"use-system-user"
            | b"user-name"
            | b"visible"
    )
}

fn reserved_prefix(raw_name: &[u8]) -> bool {
    raw_name
        .iter()
        .position(|byte| *byte == b':')
        .is_some_and(|index| matches!(&raw_name[..index], b"db" | b"office" | b"xlink"))
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE) {
        NamespaceKind::Office
    } else if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == DATABASE_NAMESPACE)
    {
        NamespaceKind::Database
    } else if matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == XLINK_NAMESPACE) {
        NamespaceKind::Xlink
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
            let value = decode_attribute_value(reader, &attribute)?;
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
            let value = decode_attribute_value(reader, &attribute)?;
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

fn decode_attribute_value(reader: &NsReader<&[u8]>, attribute: &Attribute<'_>) -> Result<String> {
    let decoded = reader
        .decoder()
        .decode(attribute.value.as_ref())
        .map_err(|error| invalid(&format!("invalid ODB attribute value: {error}")))?;
    quick_xml::escape::unescape(decoded.as_ref())
        .map(|value| value.into_owned())
        .map_err(|error| invalid(&format!("invalid ODB attribute value: {error}")))
}

fn parse_bool(value: &str) -> Result<bool> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(invalid("invalid ODB boolean attribute")),
    }
}

fn parse_positive_u64(value: &str, kind: &str) -> Result<u64> {
    let value = collapse_xml_whitespace(value)?
        .parse::<u64>()
        .map_err(|_| invalid(&format!("invalid ODB {kind}")))?;
    if value == 0 {
        return Err(invalid(&format!("ODB {kind} must be positive")));
    }
    Ok(value)
}

fn parse_namespaced_token(
    reader: &NsReader<&[u8]>,
    value: &str,
    kind: &str,
) -> Result<Option<String>> {
    let value = collapse_xml_whitespace(value)?;
    let mut parts = value.split(':');
    let prefix = parts.next().unwrap_or_default();
    let local = parts.next().unwrap_or_default();
    if !is_ncname(prefix) || !is_ncname(local) || parts.next().is_some() {
        return Err(invalid(&format!("invalid ODB {kind}")));
    }
    match reader
        .resolver()
        .resolve_prefix(QName(value.as_bytes()).prefix(), false)
    {
        ResolveResult::Bound(Namespace(uri)) => std::str::from_utf8(uri)
            .map(str::to_owned)
            .map(Some)
            .map_err(|_| invalid(&format!("invalid ODB {kind} namespace"))),
        ResolveResult::Unknown(_) | ResolveResult::Unbound => {
            Err(invalid(&format!("ODB {kind} prefix is not bound")))
        },
    }
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first == '_' || first.is_alphabetic()) {
        return false;
    }
    chars.all(|character| {
        character == '_'
            || character == '-'
            || character == '.'
            || character.is_alphanumeric()
            || character == '\u{b7}'
            || matches!(character as u32, 0x0300..=0x036f | 0x203f..=0x2040)
    })
}

fn collapse_xml_whitespace(value: &str) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve(value.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB semantic token",
            source,
        })?;
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\n' | '\r') {
            if !output.is_empty() {
                pending_space = true;
            }
        } else {
            if pending_space {
                output.push(' ');
                pending_space = false;
            }
            output.push(character);
        }
    }
    Ok(output)
}

fn parse_key_kind(value: &str) -> Result<KeyKind> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "primary" => Ok(KeyKind::Primary),
        "unique" => Ok(KeyKind::Unique),
        "foreign" => Ok(KeyKind::Foreign),
        _ => Err(invalid("invalid ODB key type")),
    }
}

fn parse_referential_action(value: &str) -> Result<ReferentialAction> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "cascade" => Ok(ReferentialAction::Cascade),
        "restrict" => Ok(ReferentialAction::Restrict),
        "set-null" => Ok(ReferentialAction::SetNull),
        "no-action" => Ok(ReferentialAction::NoAction),
        "set-default" => Ok(ReferentialAction::SetDefault),
        _ => Err(invalid("invalid ODB referential action")),
    }
}

fn parse_nullability(value: &str) -> Result<bool> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
        "nullable" => Ok(true),
        "no-nulls" => Ok(false),
        _ => Err(invalid("invalid ODB column nullability")),
    }
}

fn parse_positive_integer(value: &str, kind: &str) -> Result<u64> {
    let value = collapse_xml_whitespace(value)?;
    let parsed = value
        .parse::<u64>()
        .map_err(|_error| invalid(&format!("invalid ODB column {kind}")))?;
    if parsed == 0 {
        return Err(invalid(&format!("ODB column {kind} must be positive")));
    }
    Ok(parsed)
}

fn validate_data_type(value: &str) -> Result<DataType> {
    let value = collapse_xml_whitespace(value)?;
    match value.as_str() {
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

#[cfg(test)]
mod tests {
    use super::{Catalog, Limits};
    use crate::model::{parse_count, parse_test_lock, reset_parse_count};

    const SOURCE: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">"#,
        r#"<office:body><office:database><db:data-source><db:connection-data><db:connection-resource xlink:href="" xlink:type="simple"/></db:connection-data></db:data-source></office:database></office:body>"#,
        r#"</office:document-content>"#,
    );
    const INVALID_SETTINGS_SOURCE: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:xlink="http://www.w3.org/1999/xlink">"#,
        r#"<office:body><office:database><db:data-source><db:connection-data>"#,
        r#"<db:connection-resource xlink:href="" xlink:type="simple"/><db:login db:login-timeout="not-a-number"/></db:connection-data>"#,
        r#"</db:data-source></office:database></office:body></office:document-content>"#,
    );

    #[test]
    fn settings_projection_is_lazy_and_reused() {
        let _lock = parse_test_lock().lock().unwrap();
        reset_parse_count();
        let catalog = Catalog::parse(SOURCE, Limits::default()).unwrap();
        assert_eq!(parse_count(), 0);
        assert!(catalog.tables().is_empty());
        assert_eq!(parse_count(), 0);

        let first = catalog.settings().unwrap();
        assert!(first.is_empty());
        assert_eq!(parse_count(), 1);
        let second = catalog.settings().unwrap();
        assert!(std::ptr::eq(first, second));
        assert_eq!(parse_count(), 1);
    }

    #[test]
    fn failed_settings_reads_are_not_cached() {
        let _lock = parse_test_lock().lock().unwrap();
        reset_parse_count();
        let catalog = Catalog::parse(INVALID_SETTINGS_SOURCE, Limits::default()).unwrap();
        assert!(catalog.settings().is_err());
        assert!(catalog.settings().is_err());
        assert_eq!(parse_count(), 2);
    }
}
