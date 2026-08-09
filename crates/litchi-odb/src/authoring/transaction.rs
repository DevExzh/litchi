//! Unified, source-checked ODB package transactions.
#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cognitive_complexity,
    clippy::missing_errors_doc,
    clippy::similar_names,
    clippy::too_many_arguments,
    clippy::too_many_lines,
    reason = "the unified transaction keeps related public CRUD verbs and its bounded XML splice engine together"
)]

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::ops::Range;

use crate::{
    Column, Component, ComponentKind, Connection, Database, Index, Key, Query, ReferentialAction,
    Table, TableKind,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_VALUE_BYTES: usize = 1024 * 1024;
const MAX_OPERATIONS: usize = 65_536;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;

/// The semantic family changed by one staged operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChangeKind {
    Connection,
    Query,
    Table,
    Column,
    Key,
    Index,
    Component,
    ProducerExtension,
}

/// One ordered semantic effect in a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Change {
    kind: ChangeKind,
    target: String,
}

impl Change {
    /// Returns the semantic family affected by this operation.
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        self.kind
    }

    /// Returns the stable producer-visible target key.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }
}

/// A source-bound unified edit over one immutable database snapshot.
///
/// All values remain inert XML metadata. No method opens a connection, loads a
/// driver, follows a component link, or executes a stored command.
pub struct Edit<'source> {
    source: &'source Database,
    content: String,
    changes: Vec<Change>,
    legacy_query: Option<QueryChange>,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Database) -> Self {
        Self {
            source,
            content: source.content_xml().to_owned(),
            changes: Vec::new(),
            legacy_query: None,
        }
    }

    /// Returns the currently staged XML without publishing a package.
    #[must_use]
    pub fn staged_content_xml(&self) -> &str {
        &self.content
    }

    /// Returns semantic effects in call order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Adds a table presentation or schema definition at the collection tail.
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        validate_name(table.name(), "table")?;
        if find_named_table(&scan(&self.content)?, table.name())?.is_some() {
            return invalid("ODB table already exists");
        }
        let prefixes = prefixes(&self.content)?;
        let fragment = serialize_table(&table, &prefixes.database);
        let nodes = scan(&self.content)?;
        match table.kind() {
            TableKind::Representation => insert_into_or_create_collection(
                &mut self.content,
                &nodes,
                "table-representations",
                "database",
                &fragment,
                &prefixes.database,
            )?,
            TableKind::Definition => {
                if let Some(collection) = unique_node(&nodes, "table-definitions")? {
                    insert_child(&mut self.content, collection, &fragment)?;
                } else if let Some(schema) = unique_node(&nodes, "schema-definition")? {
                    let wrapped = format!(
                        "<{0}:table-definitions>{fragment}</{0}:table-definitions>",
                        prefixes.database
                    );
                    insert_child(&mut self.content, schema, &wrapped)?;
                } else {
                    let database = unique_required_node(&nodes, "database")?;
                    let wrapped = format!(
                        "<{0}:schema-definition><{0}:table-definitions>{fragment}</{0}:table-definitions></{0}:schema-definition>",
                        prefixes.database
                    );
                    insert_child(&mut self.content, database, &wrapped)?;
                }
            },
        }
        self.record(ChangeKind::Table, table.name())
    }

    /// Replaces one complete table declaration while preserving adjacent XML.
    pub fn replace_table(&mut self, name: &str, table: Table) -> Result<()> {
        validate_name(table.name(), "table")?;
        let nodes = scan(&self.content)?;
        let site = find_named_table(&nodes, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        if site.local != table_element(table.kind()) {
            return invalid("ODB table replacement cannot change declaration kind");
        }
        if table.name() != name {
            return Err(Error::Unsupported(
                "ODB table replacement cannot rename; use rename_table first".to_string(),
            ));
        }
        let prefix = prefixes(&self.content)?.database;
        replace_span(
            &mut self.content,
            site.full.clone(),
            &serialize_table(&table, &prefix),
        )?;
        self.record(ChangeKind::Table, name)
    }

    /// Removes one unambiguous table declaration.
    pub fn remove_table(&mut self, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_named_table(&nodes, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        if nodes.iter().any(|node| {
            node.local == "key"
                && attribute(node, DATABASE_NAMESPACE, "referenced-table-name")
                    .is_some_and(|value| value == name)
                && !ancestors(&nodes, node).any(|ancestor| ancestor.id == site.id)
        }) {
            return Err(Error::Unsupported(
                "ODB table removal would orphan an incoming relation".to_string(),
            ));
        }
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(ChangeKind::Table, name)
    }

    /// Renames a table and all modeled foreign-key table references atomically.
    pub fn rename_table(&mut self, name: &str, replacement: &str) -> Result<()> {
        validate_name(replacement, "table")?;
        let nodes = scan(&self.content)?;
        if find_named_table(&nodes, replacement)?.is_some() {
            return invalid("ODB replacement table name already exists");
        }
        let site = find_named_table(&nodes, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        let db_prefix = prefixes(&self.content)?.database;
        let mut edits = vec![attribute_edit(
            &self.content,
            site,
            DATABASE_NAMESPACE,
            "name",
            Some(replacement),
            &db_prefix,
        )?];
        for key in nodes.iter().filter(|node| node.local == "key") {
            if attribute(key, DATABASE_NAMESPACE, "referenced-table-name")
                .is_some_and(|value| value == name)
            {
                edits.push(attribute_edit(
                    &self.content,
                    key,
                    DATABASE_NAMESPACE,
                    "referenced-table-name",
                    Some(replacement),
                    &db_prefix,
                )?);
            }
        }
        apply_edits(&mut self.content, edits)?;
        self.record(ChangeKind::Table, name)
    }

    /// Adds a column to a table at the collection tail.
    pub fn add_column(&mut self, table: &str, column: Column) -> Result<()> {
        validate_name(column.name(), "column")?;
        let nodes = scan(&self.content)?;
        let table_site = find_named_table(&nodes, table)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        if find_named_child(
            &nodes,
            table_site,
            column_element(table_site),
            column.name(),
        )?
        .is_some()
        {
            return invalid("ODB column already exists");
        }
        let prefix = prefixes(&self.content)?.database;
        let fragment = serialize_column(&column, table_site.local == "table-definition", &prefix);
        insert_table_collection(
            &mut self.content,
            &nodes,
            table_site,
            column_collection(table_site),
            &fragment,
            &prefix,
        )?;
        self.record(ChangeKind::Column, &format!("{table}/{}", column.name()))
    }

    /// Replaces one complete column declaration.
    pub fn replace_column(&mut self, table: &str, name: &str, column: Column) -> Result<()> {
        validate_name(column.name(), "column")?;
        let nodes = scan(&self.content)?;
        let table_site = find_named_table(&nodes, table)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        let site = find_named_child(&nodes, table_site, column_element(table_site), name)?
            .ok_or_else(|| Error::InvalidFormat("ODB column does not exist".to_string()))?;
        if column.name() != name {
            return Err(Error::Unsupported(
                "ODB column replacement cannot rename; use rename_column first".to_string(),
            ));
        }
        let prefix = prefixes(&self.content)?.database;
        let fragment = serialize_column(&column, table_site.local == "table-definition", &prefix);
        replace_span(&mut self.content, site.full.clone(), &fragment)?;
        self.record(ChangeKind::Column, &format!("{table}/{name}"))
    }

    /// Removes one table column.
    pub fn remove_column(&mut self, table: &str, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let table_site = find_named_table(&nodes, table)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        let site = find_named_child(&nodes, table_site, column_element(table_site), name)?
            .ok_or_else(|| Error::InvalidFormat("ODB column does not exist".to_string()))?;
        if column_is_referenced(&nodes, table_site, name) {
            return Err(Error::Unsupported(
                "ODB column removal would orphan a key or index reference".to_string(),
            ));
        }
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(ChangeKind::Column, &format!("{table}/{name}"))
    }

    /// Renames a column and local key/index mappings atomically.
    pub fn rename_column(&mut self, table: &str, name: &str, replacement: &str) -> Result<()> {
        validate_name(replacement, "column")?;
        let nodes = scan(&self.content)?;
        let table_site = find_named_table(&nodes, table)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        if find_named_child(&nodes, table_site, column_element(table_site), replacement)?.is_some()
        {
            return invalid("ODB replacement column name already exists");
        }
        let site = find_named_child(&nodes, table_site, column_element(table_site), name)?
            .ok_or_else(|| Error::InvalidFormat("ODB column does not exist".to_string()))?;
        let db_prefix = prefixes(&self.content)?.database;
        let mut edits = vec![attribute_edit(
            &self.content,
            site,
            DATABASE_NAMESPACE,
            "name",
            Some(replacement),
            &db_prefix,
        )?];
        for node in descendants(&nodes, table_site).filter(|node| {
            matches!(node.local.as_str(), "key-column" | "index-column")
                && attribute(node, DATABASE_NAMESPACE, "name").is_some_and(|value| value == name)
        }) {
            edits.push(attribute_edit(
                &self.content,
                node,
                DATABASE_NAMESPACE,
                "name",
                Some(replacement),
                &db_prefix,
            )?);
        }
        for key in nodes.iter().filter(|node| {
            node.local == "key"
                && attribute(node, DATABASE_NAMESPACE, "referenced-table-name")
                    .is_some_and(|value| value == table)
        }) {
            for node in descendants(&nodes, key).filter(|node| {
                node.local == "key-column"
                    && attribute(node, DATABASE_NAMESPACE, "related-column-name")
                        .is_some_and(|value| value == name)
            }) {
                edits.push(attribute_edit(
                    &self.content,
                    node,
                    DATABASE_NAMESPACE,
                    "related-column-name",
                    Some(replacement),
                    &db_prefix,
                )?);
            }
        }
        apply_edits(&mut self.content, edits)?;
        self.record(ChangeKind::Column, &format!("{table}/{name}"))
    }

    /// Adds a schema key. Foreign keys are the ODF relation representation.
    pub fn add_key(&mut self, table: &str, key: Key) -> Result<()> {
        let name = key
            .name()
            .ok_or_else(|| Error::InvalidFormat("ODB authored key requires a name".to_string()))?;
        validate_name(name, "key")?;
        let prefix = prefixes(&self.content)?.database;
        self.add_table_child(
            table,
            "key",
            name,
            "keys",
            &serialize_key(&key, &prefix),
            ChangeKind::Key,
        )
    }

    /// Replaces a schema key, including its complete relation mapping.
    pub fn replace_key(&mut self, table: &str, name: &str, key: Key) -> Result<()> {
        let replacement_name = key
            .name()
            .ok_or_else(|| Error::InvalidFormat("ODB authored key requires a name".to_string()))?;
        validate_name(replacement_name, "key")?;
        let prefix = prefixes(&self.content)?.database;
        self.replace_table_child(
            table,
            "key",
            name,
            replacement_name,
            &serialize_key(&key, &prefix),
            ChangeKind::Key,
        )
    }

    /// Removes a schema key or relation.
    pub fn remove_key(&mut self, table: &str, name: &str) -> Result<()> {
        self.remove_table_child(table, "key", name)
    }

    /// Adds an index to a schema table.
    pub fn add_index(&mut self, table: &str, index: Index) -> Result<()> {
        validate_name(index.name(), "index")?;
        let prefix = prefixes(&self.content)?.database;
        self.add_table_child(
            table,
            "index",
            index.name(),
            "indices",
            &serialize_index(&index, &prefix),
            ChangeKind::Index,
        )
    }

    /// Replaces an index declaration.
    pub fn replace_index(&mut self, table: &str, name: &str, index: Index) -> Result<()> {
        validate_name(index.name(), "index")?;
        let prefix = prefixes(&self.content)?.database;
        self.replace_table_child(
            table,
            "index",
            name,
            index.name(),
            &serialize_index(&index, &prefix),
            ChangeKind::Index,
        )
    }

    /// Removes an index declaration.
    pub fn remove_index(&mut self, table: &str, name: &str) -> Result<()> {
        self.remove_table_child(table, "index", name)
    }

    /// Adds a stored query without interpreting its command.
    pub fn add_query(&mut self, query: Query) -> Result<()> {
        validate_name(query.name(), "query")?;
        validate_value(query.command(), "query command")?;
        let nodes = scan(&self.content)?;
        if find_named_node(&nodes, "query", query.name())?.is_some() {
            return invalid("ODB query already exists");
        }
        let prefix = prefixes(&self.content)?.database;
        let fragment = serialize_query(&query, &prefix);
        insert_into_or_create_collection(
            &mut self.content,
            &nodes,
            "queries",
            "database",
            &fragment,
            &prefix,
        )?;
        self.record(ChangeKind::Query, query.name())
    }

    /// Replaces a stored query without interpreting its command.
    pub fn replace_query(&mut self, name: &str, query: Query) -> Result<()> {
        validate_name(query.name(), "query")?;
        validate_value(query.command(), "query command")?;
        let nodes = scan(&self.content)?;
        let site = find_named_node(&nodes, "query", name)?
            .ok_or_else(|| Error::InvalidFormat("ODB query does not exist".to_string()))?;
        if query.name() != name && find_named_node(&nodes, "query", query.name())?.is_some() {
            return invalid("ODB replacement query name already exists");
        }
        let prefix = prefixes(&self.content)?.database;
        replace_span(
            &mut self.content,
            site.full.clone(),
            &serialize_query(&query, &prefix),
        )?;
        self.record(ChangeKind::Query, name)
    }

    /// Removes a stored query.
    pub fn remove_query(&mut self, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_named_node(&nodes, "query", name)?
            .ok_or_else(|| Error::InvalidFormat("ODB query does not exist".to_string()))?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(ChangeKind::Query, name)
    }

    /// Replaces the inert command text stored for one exactly named query.
    pub fn set_query_command(&mut self, name: &str, value: impl Into<String>) -> Result<()> {
        let value = value.into();
        validate_value(&value, "query command")?;
        self.prepare_legacy_query(name)?;
        self.set_named_db_attribute("query", name, "command", Some(&value))?;
        if let Some(change) = self.legacy_query.as_mut() {
            change.after_command = value;
        }
        self.record(ChangeKind::Query, name)
    }

    /// Sets or removes the stored `db:escape-processing` declaration.
    pub fn set_query_escape_processing(&mut self, name: &str, value: Option<bool>) -> Result<()> {
        self.prepare_legacy_query(name)?;
        let lexical = value.map(bool_text);
        self.set_named_db_attribute("query", name, "escape-processing", lexical)?;
        if let Some(change) = self.legacy_query.as_mut() {
            change.after_escape_processing = value;
        }
        self.record(ChangeKind::Query, name)
    }

    /// Replaces, creates, or removes the inert database connection target.
    pub fn set_connection(&mut self, connection: Option<Connection>) -> Result<()> {
        let nodes = scan(&self.content)?;
        let prefixes = prefixes(&self.content)?;
        let targets = nodes
            .iter()
            .filter(|node| {
                node.namespace == NamespaceKind::Database
                    && matches!(
                        node.local.as_str(),
                        "connection-resource" | "file-based-database" | "server-database"
                    )
            })
            .collect::<Vec<_>>();
        if targets.len() > 1 {
            return invalid("ODB connection target is ambiguous");
        }
        match (targets.first().copied(), connection.as_ref()) {
            (Some(site), Some(value)) => {
                let owner = connection_owner(&nodes, site);
                replace_span(
                    &mut self.content,
                    owner.full.clone(),
                    &serialize_connection_owner(value, &prefixes),
                )?;
            },
            (Some(site), None) => {
                let owner = connection_owner(&nodes, site);
                replace_span(&mut self.content, owner.full.clone(), "")?;
            },
            (None, Some(value)) => {
                let fragment = serialize_connection(value, &prefixes);
                if let Some(description) = unique_node(&nodes, "database-description")? {
                    insert_child(&mut self.content, description, &fragment)?;
                } else if let Some(data) = unique_node(&nodes, "connection-data")? {
                    insert_child(
                        &mut self.content,
                        data,
                        &serialize_connection_owner(value, &prefixes),
                    )?;
                } else {
                    let source = unique_required_node(&nodes, "data-source")?;
                    let wrapped = format!(
                        "<{0}:connection-data>{1}</{0}:connection-data>",
                        prefixes.database,
                        serialize_connection_owner(value, &prefixes)
                    );
                    insert_child(&mut self.content, source, &wrapped)?;
                }
            },
            (None, None) => return Ok(()),
        }
        self.record(ChangeKind::Connection, "connection")
    }

    /// Adds an inert form or report component.
    pub fn add_component(&mut self, component: Component) -> Result<()> {
        let name = component.name().ok_or_else(|| {
            Error::InvalidFormat("ODB authored component requires a name".to_string())
        })?;
        validate_name(name, "component")?;
        let nodes = scan(&self.content)?;
        if find_component(&nodes, component.kind(), name)?.is_some() {
            return invalid("ODB component already exists");
        }
        let prefix = prefixes(&self.content)?.database;
        let collection = component_collection(component.kind());
        let prefixes = prefixes(&self.content)?;
        let fragment = serialize_component(&component, &prefixes);
        insert_into_or_create_collection(
            &mut self.content,
            &nodes,
            collection,
            "database",
            &fragment,
            &prefix,
        )?;
        self.record(ChangeKind::Component, &format!("{collection}/{name}"))
    }

    /// Replaces an inert form or report component.
    pub fn replace_component(
        &mut self,
        kind: ComponentKind,
        name: &str,
        component: Component,
    ) -> Result<()> {
        if component.kind() != kind {
            return invalid("ODB component replacement cannot change its collection kind");
        }
        let replacement_name = component.name().ok_or_else(|| {
            Error::InvalidFormat("ODB authored component requires a name".to_string())
        })?;
        validate_name(replacement_name, "component")?;
        let nodes = scan(&self.content)?;
        let site = find_component(&nodes, kind, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB component does not exist".to_string()))?;
        if replacement_name != name && find_component(&nodes, kind, replacement_name)?.is_some() {
            return invalid("ODB replacement component name already exists");
        }
        let prefixes = prefixes(&self.content)?;
        replace_span(
            &mut self.content,
            site.full.clone(),
            &serialize_component(&component, &prefixes),
        )?;
        self.record(
            ChangeKind::Component,
            &format!("{}/{name}", component_collection(kind)),
        )
    }

    /// Removes an inert form or report component.
    pub fn remove_component(&mut self, kind: ComponentKind, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_component(&nodes, kind, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB component does not exist".to_string()))?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(
            ChangeKind::Component,
            &format!("{}/{name}", component_collection(kind)),
        )
    }

    /// Adds one compact producer-extension subtree below `office:database`.
    ///
    /// The extension is preserved as inert XML and must use a namespace other
    /// than the ODF office and database namespaces.
    pub fn add_producer_extension(&mut self, xml: &str) -> Result<()> {
        validate_extension(xml)?;
        let extension_nodes = scan(xml)?;
        let root = extension_nodes
            .iter()
            .find(|node| node.parent.is_none())
            .ok_or_else(|| {
                Error::InvalidFormat("ODB producer extension has no root".to_string())
            })?;
        if matches!(
            root.namespace,
            NamespaceKind::Office | NamespaceKind::Database
        ) {
            return invalid("ODB producer extension must use a producer namespace");
        }
        let nodes = scan(&self.content)?;
        let database = unique_required_node(&nodes, "database")?;
        insert_child(&mut self.content, database, xml)?;
        self.record(ChangeKind::ProducerExtension, &root.local)
    }

    /// Removes one unambiguous direct producer-extension child by namespace URI
    /// and local name.
    pub fn remove_producer_extension(&mut self, namespace: &str, local: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let database = unique_required_node(&nodes, "database")?;
        let matches = nodes
            .iter()
            .filter(|node| {
                node.parent == Some(database.id)
                    && node.namespace_uri.as_deref() == Some(namespace)
                    && node.local == local
                    && node.namespace == NamespaceKind::Other
            })
            .collect::<Vec<_>>();
        let site = unique_match(matches, "ODB producer extension selector is ambiguous")?
            .ok_or_else(|| {
                Error::InvalidFormat("ODB producer extension does not exist".to_string())
            })?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(ChangeKind::ProducerExtension, local)
    }

    fn add_table_child(
        &mut self,
        table: &str,
        local: &str,
        name: &str,
        collection: &str,
        fragment: &str,
        kind: ChangeKind,
    ) -> Result<()> {
        let nodes = scan(&self.content)?;
        let table_site = definition_table(&nodes, table)?;
        if find_named_child(&nodes, table_site, local, name)?.is_some() {
            return invalid("ODB table child already exists");
        }
        let prefix = prefixes(&self.content)?.database;
        insert_table_collection(
            &mut self.content,
            &nodes,
            table_site,
            collection,
            fragment,
            &prefix,
        )?;
        self.record(kind, &format!("{table}/{name}"))
    }

    fn replace_table_child(
        &mut self,
        table: &str,
        local: &str,
        name: &str,
        replacement_name: &str,
        fragment: &str,
        kind: ChangeKind,
    ) -> Result<()> {
        let nodes = scan(&self.content)?;
        let table_site = definition_table(&nodes, table)?;
        let site = find_named_child(&nodes, table_site, local, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table child does not exist".to_string()))?;
        if replacement_name != name
            && find_named_child(&nodes, table_site, local, replacement_name)?.is_some()
        {
            return invalid("ODB replacement table child name already exists");
        }
        replace_span(&mut self.content, site.full.clone(), fragment)?;
        self.record(kind, &format!("{table}/{name}"))
    }

    fn remove_table_child(&mut self, table: &str, local: &str, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let table_site = definition_table(&nodes, table)?;
        let site = find_named_child(&nodes, table_site, local, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table child does not exist".to_string()))?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        let kind = if local == "key" {
            ChangeKind::Key
        } else {
            ChangeKind::Index
        };
        self.record(kind, &format!("{table}/{name}"))
    }

    fn set_named_db_attribute(
        &mut self,
        local: &str,
        name: &str,
        attribute_name: &str,
        value: Option<&str>,
    ) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_named_node(&nodes, local, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB edit selector did not match".to_string()))?;
        let prefix = prefixes(&self.content)?.database;
        let edit = attribute_edit(
            &self.content,
            site,
            DATABASE_NAMESPACE,
            attribute_name,
            value,
            &prefix,
        )?;
        apply_edits(&mut self.content, vec![edit])
    }

    fn prepare_legacy_query(&mut self, name: &str) -> Result<()> {
        if self
            .legacy_query
            .as_ref()
            .is_some_and(|change| change.name != name)
        {
            return invalid("legacy query scalar setters support one query per transaction");
        }
        if self.legacy_query.is_none() {
            let catalog = self.source.catalog()?;
            let query = catalog.query(name)?.ok_or_else(|| {
                Error::InvalidFormat(format!("ODB query '{name}' does not exist"))
            })?;
            self.legacy_query = Some(QueryChange {
                name: name.to_owned(),
                before_command: query.command().to_owned(),
                after_command: query.command().to_owned(),
                before_escape_processing: query.escape_processing(),
                after_escape_processing: query.escape_processing(),
            });
        }
        Ok(())
    }

    fn record(&mut self, kind: ChangeKind, target: &str) -> Result<()> {
        if self.changes.len() >= MAX_OPERATIONS {
            return invalid("ODB transaction exceeds the operation limit");
        }
        self.changes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB transaction operations",
                source,
            })?;
        self.changes.push(Change {
            kind,
            target: target.to_owned(),
        });
        Ok(())
    }

    /// Atomically rebuilds, fully reopens, and semantically verifies the
    /// candidate package.
    pub fn commit(self) -> Result<Commit> {
        if self.content == self.source.content_xml() {
            return Ok(Commit::unchanged(self.source.clone()));
        }
        crate::codec::validate(&self.content)?;
        let snapshot = Database {
            package: self.source.package.rebuild_with_content(&self.content)?,
        };
        snapshot.catalog()?;
        let legacy_query = self.legacy_query.filter(|change| !change.is_noop());
        Ok(Commit {
            patch: Patch {
                source: self.source.clone(),
                target: snapshot.clone(),
                changes: self.changes,
                legacy_query,
            },
            snapshot,
            changed: true,
        })
    }
}

/// One reversible stored-query scalar operation retained for compatibility.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct QueryChange {
    name: String,
    before_command: String,
    after_command: String,
    before_escape_processing: Option<bool>,
    after_escape_processing: Option<bool>,
}

impl QueryChange {
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    #[must_use]
    pub fn before_command(&self) -> &str {
        &self.before_command
    }

    #[must_use]
    pub fn after_command(&self) -> &str {
        &self.after_command
    }

    #[must_use]
    pub const fn before_escape_processing(&self) -> Option<bool> {
        self.before_escape_processing
    }

    #[must_use]
    pub const fn after_escape_processing(&self) -> Option<bool> {
        self.after_escape_processing
    }

    fn is_noop(&self) -> bool {
        self.before_escape_processing == self.after_escape_processing
            && self.before_command == self.after_command
    }
}

/// A committed immutable database and its source-checked reversible patch.
pub struct Commit {
    snapshot: Database,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Database) -> Self {
        Self {
            patch: Patch {
                source: snapshot.clone(),
                target: snapshot.clone(),
                changes: Vec::new(),
                legacy_query: None,
            },
            snapshot,
            changed: false,
        }
    }

    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn database(&self) -> &Database {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub fn into_database(self) -> Database {
        self.snapshot
    }
}

/// A byte-exact source-checked reversible ODB patch.
#[derive(Clone)]
pub struct Patch {
    source: Database,
    target: Database,
    changes: Vec<Change>,
    legacy_query: Option<QueryChange>,
}

impl Patch {
    #[must_use]
    pub fn is_applicable_to(&self, source: &Database) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    pub fn apply(&self, source: &Database) -> Result<Database> {
        if !self.is_applicable_to(source) {
            return invalid("ODB patch source does not match its expected snapshot");
        }
        Ok(self.target.clone())
    }

    /// Returns all semantic effects in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Returns the legacy scalar query change when this transaction used it.
    #[must_use]
    pub const fn change(&self) -> Option<&QueryChange> {
        self.legacy_query.as_ref()
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.changes.iter().rev().cloned().collect(),
            legacy_query: self.legacy_query.as_ref().map(|change| QueryChange {
                name: change.name.clone(),
                before_command: change.after_command.clone(),
                after_command: change.before_command.clone(),
                before_escape_processing: change.after_escape_processing,
                after_escape_processing: change.before_escape_processing,
            }),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Database,
    Xlink,
    Other,
}

#[derive(Clone)]
struct Attr {
    namespace_uri: Option<String>,
    local: String,
    qname: String,
    value: String,
}

#[derive(Clone)]
struct Node {
    id: usize,
    parent: Option<usize>,
    namespace: NamespaceKind,
    namespace_uri: Option<String>,
    local: String,
    start_tag: Range<usize>,
    full: Range<usize>,
    attrs: Vec<Attr>,
}

struct Prefixes {
    database: String,
    xlink: String,
}

struct TextEdit {
    range: Range<usize>,
    value: String,
}

fn scan(source: &str) -> Result<Vec<Node>> {
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut nodes = Vec::<Node>::new();
    let mut stack = Vec::<usize>::new();
    loop {
        let (resolved, raw_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODB edit XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let namespace_uri = namespace_uri(&resolved)?;
        let event = raw_event.into_owned();
        drop(resolved);
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidFormat("ODB edit XML position exceeds this platform".to_string())
        })?;
        let start_event = matches!(event, Event::Start(_));
        match event {
            Event::Start(element) | Event::Empty(element) => {
                let start = source[..end].rfind('<').ok_or_else(|| {
                    Error::InvalidFormat("ODB element start is missing".to_string())
                })?;
                let id = nodes.len();
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("ODB local name is not UTF-8".to_string())
                    })?
                    .to_owned();
                nodes.push(Node {
                    id,
                    parent: stack.last().copied(),
                    namespace,
                    namespace_uri,
                    local,
                    start_tag: start..end,
                    full: start..end,
                    attrs: decode_attributes(&reader, &element)?,
                });
                if start_event {
                    stack.push(id);
                }
            },
            Event::End(_) => {
                let id = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODB edit XML has an unmatched end tag".to_string())
                })?;
                nodes[id].full.end = end;
            },
            Event::DocType(_) => return invalid("DOCTYPE is not allowed in ODB edit XML"),
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
    if !stack.is_empty() {
        return invalid("ODB edit XML is incomplete");
    }
    Ok(nodes)
}

pub(crate) fn producer_extensions(source: &str) -> Result<Vec<crate::ProducerExtension>> {
    let nodes = scan(source)?;
    let database = unique_required_node(&nodes, "database")?;
    let matches = nodes
        .iter()
        .filter(|node| {
            node.parent == Some(database.id)
                && node.namespace == NamespaceKind::Other
                && node.namespace_uri.is_some()
        })
        .collect::<Vec<_>>();
    let mut extensions = Vec::new();
    extensions
        .try_reserve_exact(matches.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB producer-extension catalog",
            source,
        })?;
    for node in matches {
        let namespace = node.namespace_uri.clone().ok_or_else(|| {
            Error::InvalidFormat("ODB producer-extension namespace is missing".to_string())
        })?;
        let xml = source
            .get(node.full.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("ODB producer-extension span is invalid".to_string())
            })?
            .to_owned();
        extensions.push(crate::ProducerExtension {
            namespace,
            local_name: node.local.clone(),
            xml,
        });
    }
    Ok(extensions)
}

fn decode_attributes(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Vec<Attr>> {
    let mut attributes = Vec::new();
    for raw in element.attributes() {
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid ODB attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let qname = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|_error| Error::InvalidFormat("ODB attribute name is not UTF-8".to_string()))?
            .to_owned();
        let local = std::str::from_utf8(local.as_ref())
            .map_err(|_error| {
                Error::InvalidFormat("ODB attribute local name is not UTF-8".to_string())
            })?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map(std::borrow::Cow::into_owned)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODB attribute value: {error}"))
            })?;
        attributes.push(Attr {
            namespace_uri: namespace_uri(&resolved)?,
            local,
            qname,
            value,
        });
    }
    Ok(attributes)
}

fn namespace_kind(value: &ResolveResult<'_>) -> NamespaceKind {
    match value {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE_NAMESPACE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DATABASE_NAMESPACE => {
            NamespaceKind::Database
        },
        ResolveResult::Bound(Namespace(uri)) if *uri == XLINK_NAMESPACE => NamespaceKind::Xlink,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn namespace_uri(value: &ResolveResult<'_>) -> Result<Option<String>> {
    match value {
        ResolveResult::Bound(Namespace(uri)) => std::str::from_utf8(uri)
            .map(str::to_owned)
            .map(Some)
            .map_err(|_error| Error::InvalidFormat("ODB namespace URI is not UTF-8".to_string())),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => Ok(None),
    }
}

fn prefixes(source: &str) -> Result<Prefixes> {
    let nodes = scan(source)?;
    let database = nodes
        .iter()
        .find(|node| node.namespace == NamespaceKind::Database)
        .and_then(|node| qname_prefix(source, node));
    let database = database.ok_or_else(|| {
        Error::Unsupported("ODB editing requires a prefixed database namespace".to_string())
    })?;
    let xlink = nodes
        .iter()
        .flat_map(|node| &node.attrs)
        .find(|attribute| attribute.namespace_uri.as_deref() == Some(str_from(XLINK_NAMESPACE)))
        .and_then(|attribute| {
            attribute
                .qname
                .split_once(':')
                .map(|(prefix, _local)| prefix)
        })
        .or_else(|| {
            nodes
                .iter()
                .flat_map(|node| &node.attrs)
                .find_map(|attribute| {
                    (attribute.value == str_from(XLINK_NAMESPACE))
                        .then(|| attribute.qname.strip_prefix("xmlns:"))
                        .flatten()
                })
        })
        .unwrap_or("xlink")
        .to_owned();
    Ok(Prefixes { database, xlink })
}

fn qname_prefix(source: &str, node: &Node) -> Option<String> {
    let tag = source.get(node.start_tag.clone())?;
    let name_end = tag
        .char_indices()
        .skip(1)
        .find(|(_index, value)| value.is_ascii_whitespace() || matches!(value, '>' | '/'))
        .map_or(tag.len(), |(index, _value)| index);
    tag.get(1..name_end)?
        .split_once(':')
        .map(|(prefix, _local)| prefix.to_owned())
}

fn unique_node<'a>(nodes: &'a [Node], local: &str) -> Result<Option<&'a Node>> {
    unique_match(
        nodes
            .iter()
            .filter(|node| {
                node.local == local
                    && matches!(
                        (local, node.namespace),
                        ("database", NamespaceKind::Office) | (_, NamespaceKind::Database)
                    )
            })
            .collect(),
        "ODB edit owner is ambiguous",
    )
}

fn unique_required_node<'a>(nodes: &'a [Node], local: &str) -> Result<&'a Node> {
    unique_node(nodes, local)?
        .ok_or_else(|| Error::InvalidFormat(format!("ODB {local} owner does not exist")))
}

fn unique_match<'a>(matches: Vec<&'a Node>, message: &str) -> Result<Option<&'a Node>> {
    if matches.len() > 1 {
        return invalid(message);
    }
    Ok(matches.into_iter().next())
}

fn find_named_node<'a>(nodes: &'a [Node], local: &str, name: &str) -> Result<Option<&'a Node>> {
    unique_match(
        nodes
            .iter()
            .filter(|node| {
                node.namespace == NamespaceKind::Database
                    && node.local == local
                    && attribute(node, DATABASE_NAMESPACE, "name")
                        .is_some_and(|value| value == name)
            })
            .collect(),
        "ODB named selector is ambiguous",
    )
}

fn find_named_table<'a>(nodes: &'a [Node], name: &str) -> Result<Option<&'a Node>> {
    unique_match(
        nodes
            .iter()
            .filter(|node| {
                node.namespace == NamespaceKind::Database
                    && matches!(
                        node.local.as_str(),
                        "table-representation" | "table-definition"
                    )
                    && attribute(node, DATABASE_NAMESPACE, "name")
                        .is_some_and(|value| value == name)
            })
            .collect(),
        "ODB table selector is ambiguous",
    )
}

fn definition_table<'a>(nodes: &'a [Node], name: &str) -> Result<&'a Node> {
    let table = find_named_table(nodes, name)?
        .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
    if table.local != "table-definition" {
        return invalid("ODB keys and indices require a table definition");
    }
    Ok(table)
}

fn find_named_child<'a>(
    nodes: &'a [Node],
    owner: &'a Node,
    local: &str,
    name: &str,
) -> Result<Option<&'a Node>> {
    unique_match(
        descendants(nodes, owner)
            .filter(|node| {
                node.namespace == NamespaceKind::Database
                    && node.local == local
                    && attribute(node, DATABASE_NAMESPACE, "name")
                        .is_some_and(|value| value == name)
            })
            .collect(),
        "ODB nested selector is ambiguous",
    )
}

fn find_component<'a>(
    nodes: &'a [Node],
    kind: ComponentKind,
    name: &str,
) -> Result<Option<&'a Node>> {
    let collection = component_collection(kind);
    unique_match(
        nodes
            .iter()
            .filter(|node| {
                node.namespace == NamespaceKind::Database
                    && node.local == "component"
                    && attribute(node, DATABASE_NAMESPACE, "name")
                        .is_some_and(|value| value == name)
                    && ancestors(nodes, node).any(|ancestor| ancestor.local == collection)
            })
            .collect(),
        "ODB component selector is ambiguous",
    )
}

fn descendants<'a>(nodes: &'a [Node], owner: &'a Node) -> impl Iterator<Item = &'a Node> {
    nodes
        .iter()
        .filter(move |node| ancestors(nodes, node).any(|ancestor| ancestor.id == owner.id))
}

fn ancestors<'a>(nodes: &'a [Node], node: &'a Node) -> impl Iterator<Item = &'a Node> {
    std::iter::successors(node.parent, move |id| nodes[*id].parent).map(|id| &nodes[id])
}

fn attribute<'a>(node: &'a Node, namespace: &[u8], local: &str) -> Option<&'a str> {
    let namespace = str_from(namespace);
    node.attrs
        .iter()
        .find(|attribute| {
            attribute.namespace_uri.as_deref() == Some(namespace) && attribute.local == local
        })
        .map(|attribute| attribute.value.as_str())
}

fn insert_into_or_create_collection(
    source: &mut String,
    nodes: &[Node],
    collection: &str,
    parent: &str,
    fragment: &str,
    prefix: &str,
) -> Result<()> {
    if let Some(owner) = unique_node(nodes, collection)? {
        insert_child(source, owner, fragment)
    } else {
        let owner = unique_required_node(nodes, parent)?;
        let wrapped = format!("<{prefix}:{collection}>{fragment}</{prefix}:{collection}>");
        insert_child(source, owner, &wrapped)
    }
}

fn insert_table_collection(
    source: &mut String,
    nodes: &[Node],
    table: &Node,
    collection: &str,
    fragment: &str,
    prefix: &str,
) -> Result<()> {
    let owners = descendants(nodes, table)
        .filter(|node| node.local == collection && node.namespace == NamespaceKind::Database)
        .collect::<Vec<_>>();
    if let Some(owner) = unique_match(owners, "ODB table collection is ambiguous")? {
        insert_child(source, owner, fragment)
    } else {
        let wrapped = format!("<{prefix}:{collection}>{fragment}</{prefix}:{collection}>");
        insert_child(source, table, &wrapped)
    }
}

fn insert_child(source: &mut String, owner: &Node, fragment: &str) -> Result<()> {
    if owner.full == owner.start_tag {
        let raw = source
            .get(owner.start_tag.clone())
            .ok_or_else(|| Error::InvalidFormat("ODB empty owner span is invalid".to_string()))?;
        let head = raw
            .strip_suffix("/>")
            .ok_or_else(|| Error::InvalidFormat("ODB empty owner syntax is invalid".to_string()))?;
        let name = raw
            .get(1..)
            .and_then(|value| {
                value
                    .split(|character: char| {
                        character.is_ascii_whitespace() || matches!(character, '/' | '>')
                    })
                    .next()
            })
            .ok_or_else(|| Error::InvalidFormat("ODB empty owner name is missing".to_string()))?;
        let replacement = format!("{head}>{fragment}</{name}>");
        replace_span(source, owner.full.clone(), &replacement)
    } else {
        let closing_start = source[..owner.full.end]
            .rfind("</")
            .ok_or_else(|| Error::InvalidFormat("ODB owner closing tag is missing".to_string()))?;
        replace_span(source, closing_start..closing_start, fragment)
    }
}

fn attribute_edit(
    source: &str,
    node: &Node,
    namespace: &[u8],
    local: &str,
    value: Option<&str>,
    prefix: &str,
) -> Result<TextEdit> {
    let raw = source
        .get(node.start_tag.clone())
        .ok_or_else(|| Error::InvalidFormat("ODB start-tag span is invalid".to_string()))?;
    let existing = node.attrs.iter().find(|attribute| {
        attribute.namespace_uri.as_deref() == Some(str_from(namespace)) && attribute.local == local
    });
    let replacement = match (existing, value) {
        (Some(attribute), Some(value)) => {
            replace_attribute(raw, &attribute.qname, &xml_escape(value))?
        },
        (Some(attribute), None) => remove_attribute(raw, &attribute.qname)?,
        (None, Some(value)) => {
            insert_attribute(raw, &format!("{prefix}:{local}"), &xml_escape(value))?
        },
        (None, None) => raw.to_owned(),
    };
    Ok(TextEdit {
        range: node.start_tag.clone(),
        value: replacement,
    })
}

fn apply_edits(source: &mut String, mut edits: Vec<TextEdit>) -> Result<()> {
    edits.sort_by(|left, right| right.range.start.cmp(&left.range.start));
    let mut previous = source.len();
    for edit in edits {
        if edit.range.end > previous {
            return invalid("ODB staged XML edits overlap");
        }
        previous = edit.range.start;
        replace_span(source, edit.range, &edit.value)?;
    }
    Ok(())
}

fn replace_span(source: &mut String, range: Range<usize>, value: &str) -> Result<()> {
    if source.get(range.clone()).is_none() {
        return invalid("ODB XML edit span is invalid");
    }
    let output_size = source
        .len()
        .checked_sub(range.end - range.start)
        .and_then(|size| size.checked_add(value.len()))
        .ok_or_else(|| Error::InvalidFormat("ODB edited content size overflow".to_string()))?;
    if output_size > MAX_OUTPUT_BYTES {
        return invalid("ODB edited content exceeds the output limit");
    }
    source.replace_range(range, value);
    Ok(())
}

fn replace_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let (_, span) = find_attribute(tag, name)?
        .ok_or_else(|| Error::InvalidFormat("ODB attribute disappeared".to_string()))?;
    Ok(format!(
        "{}{}{}",
        &tag[..span.start],
        value,
        &tag[span.end..]
    ))
}

fn remove_attribute(tag: &str, name: &str) -> Result<String> {
    let (span, _) = find_attribute(tag, name)?
        .ok_or_else(|| Error::InvalidFormat("ODB attribute disappeared".to_string()))?;
    Ok(format!("{}{}", &tag[..span.start], &tag[span.end..]))
}

fn insert_attribute(tag: &str, name: &str, value: &str) -> Result<String> {
    let position = if tag.ends_with("/>") {
        tag.len() - 2
    } else if tag.ends_with('>') {
        tag.len() - 1
    } else {
        return invalid("ODB start tag has no closing delimiter");
    };
    Ok(format!(
        "{} {name}=\"{value}\"{}",
        &tag[..position],
        &tag[position..]
    ))
}

fn find_attribute(tag: &str, wanted: &str) -> Result<Option<(Range<usize>, Range<usize>)>> {
    let bytes = tag.as_bytes();
    let mut cursor = 1usize;
    while cursor < bytes.len()
        && !bytes[cursor].is_ascii_whitespace()
        && !matches!(bytes[cursor], b'>' | b'/')
    {
        cursor += 1;
    }
    while cursor < bytes.len() {
        let attribute_start = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || matches!(bytes[cursor], b'>' | b'/') {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len() && !bytes[cursor].is_ascii_whitespace() && bytes[cursor] != b'='
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() || bytes[cursor] != b'=' {
            return invalid("ODB attribute is malformed");
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| Error::InvalidFormat("ODB attribute value is missing".to_string()))?;
        if !matches!(quote, b'\'' | b'\"') {
            return invalid("ODB attribute value is not quoted");
        }
        cursor += 1;
        let value_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor == bytes.len() {
            return invalid("ODB attribute value is unterminated");
        }
        let value_end = cursor;
        cursor += 1;
        if &tag[name_start..name_end] == wanted {
            return Ok(Some((attribute_start..cursor, value_start..value_end)));
        }
    }
    Ok(None)
}

fn serialize_table(table: &Table, prefix: &str) -> String {
    let local = table_element(table.kind());
    let mut children = String::new();
    if !table.columns().is_empty() {
        let collection = if table.kind() == TableKind::Definition {
            "column-definitions"
        } else {
            "columns"
        };
        children.push_str(&format!("<{prefix}:{collection}>"));
        for column in table.columns() {
            children.push_str(&serialize_column(
                column,
                table.kind() == TableKind::Definition,
                prefix,
            ));
        }
        children.push_str(&format!("</{prefix}:{collection}>"));
    }
    if !table.keys().is_empty() {
        children.push_str(&format!("<{prefix}:keys>"));
        for key in table.keys() {
            children.push_str(&serialize_key(key, prefix));
        }
        children.push_str(&format!("</{prefix}:keys>"));
    }
    if !table.indices().is_empty() {
        children.push_str(&format!("<{prefix}:indices>"));
        for index in table.indices() {
            children.push_str(&serialize_index(index, prefix));
        }
        children.push_str(&format!("</{prefix}:indices>"));
    }
    element_with_children(
        prefix,
        local,
        [("name", table.name().to_owned())],
        &children,
    )
}

fn serialize_column(column: &Column, definition: bool, prefix: &str) -> String {
    let mut attrs = vec![("name", column.name().to_owned())];
    if let Some(value) = column.data_type() {
        attrs.push(("data-type", value.as_str().to_owned()));
    }
    if let Some(value) = column.type_name() {
        attrs.push(("type-name", value.to_owned()));
    }
    push_option(&mut attrs, "precision", column.precision());
    push_option(&mut attrs, "scale", column.scale());
    if let Some(value) = column.nullable() {
        attrs.push((
            "is-nullable",
            if value { "nullable" } else { "no-nulls" }.to_owned(),
        ));
    }
    push_bool(&mut attrs, "is-empty-allowed", column.empty_allowed());
    push_bool(&mut attrs, "is-autoincrement", column.autoincrement());
    empty_element(
        prefix,
        if definition {
            "column-definition"
        } else {
            "column"
        },
        attrs,
    )
}

fn serialize_key(key: &Key, prefix: &str) -> String {
    let mut attrs = Vec::new();
    if let Some(value) = key.name() {
        attrs.push(("name", value.to_owned()));
    }
    attrs.push((
        "type",
        match key.kind() {
            crate::KeyKind::Primary => "primary",
            crate::KeyKind::Unique => "unique",
            crate::KeyKind::Foreign => "foreign",
        }
        .to_owned(),
    ));
    if let Some(value) = key.referenced_table() {
        attrs.push(("referenced-table-name", value.to_owned()));
    }
    push_action(&mut attrs, "update-rule", key.update_rule());
    push_action(&mut attrs, "delete-rule", key.delete_rule());
    let mut children = String::new();
    if !key.columns().is_empty() {
        children.push_str(&format!("<{prefix}:key-columns>"));
        for column in key.columns() {
            let mut column_attrs = Vec::new();
            if let Some(value) = column.name() {
                column_attrs.push(("name", value.to_owned()));
            }
            if let Some(value) = column.related_column() {
                column_attrs.push(("related-column-name", value.to_owned()));
            }
            children.push_str(&empty_element(prefix, "key-column", column_attrs));
        }
        children.push_str(&format!("</{prefix}:key-columns>"));
    }
    element_with_children(prefix, "key", attrs, &children)
}

fn serialize_index(index: &Index, prefix: &str) -> String {
    let mut attrs = vec![("name", index.name().to_owned())];
    push_bool(&mut attrs, "is-unique", index.unique());
    push_bool(&mut attrs, "is-clustered", index.clustered());
    let mut children = String::new();
    if !index.columns().is_empty() {
        children.push_str(&format!("<{prefix}:index-columns>"));
        for column in index.columns() {
            let mut column_attrs = vec![("name", column.name().to_owned())];
            push_bool(&mut column_attrs, "is-ascending", column.ascending());
            children.push_str(&empty_element(prefix, "index-column", column_attrs));
        }
        children.push_str(&format!("</{prefix}:index-columns>"));
    }
    element_with_children(prefix, "index", attrs, &children)
}

fn serialize_query(query: &Query, prefix: &str) -> String {
    let mut attrs = vec![
        ("name", query.name().to_owned()),
        ("command", query.command().to_owned()),
    ];
    push_bool(&mut attrs, "escape-processing", query.escape_processing());
    empty_element(prefix, "query", attrs)
}

fn serialize_component(component: &Component, prefixes: &Prefixes) -> String {
    let mut attrs = Vec::new();
    if let Some(value) = component.name() {
        attrs.push(("name", value.to_owned()));
    }
    if let Some(value) = component.title() {
        attrs.push(("title", value.to_owned()));
    }
    if let Some(value) = component.description() {
        attrs.push(("description", value.to_owned()));
    }
    push_bool(&mut attrs, "as-template", component.as_template());
    let mut output = format!("<{}:component", prefixes.database);
    append_attrs(&mut output, &prefixes.database, attrs);
    if let Some(value) = component.href() {
        output.push_str(&format!(
            " {}:href=\"{}\" {}:type=\"simple\"",
            prefixes.xlink,
            xml_escape(value),
            prefixes.xlink
        ));
    }
    output.push_str("/>");
    output
}

fn serialize_connection(connection: &Connection, prefixes: &Prefixes) -> String {
    match connection {
        Connection::File(href) => format!(
            "<{0}:file-based-database {1}:href=\"{2}\"/>",
            prefixes.database,
            prefixes.xlink,
            xml_escape(href)
        ),
        Connection::Resource(href) => format!(
            "<{0}:connection-resource {1}:href=\"{2}\" {1}:type=\"simple\"/>",
            prefixes.database,
            prefixes.xlink,
            xml_escape(href)
        ),
        Connection::Server { host, database } => format!(
            "<{0}:server-database {0}:hostname=\"{1}\" {0}:database-name=\"{2}\"/>",
            prefixes.database,
            xml_escape(host),
            xml_escape(database)
        ),
    }
}

fn serialize_connection_owner(connection: &Connection, prefixes: &Prefixes) -> String {
    let target = serialize_connection(connection, prefixes);
    if matches!(connection, Connection::Resource(_)) {
        target
    } else {
        format!(
            "<{0}:database-description>{target}</{0}:database-description>",
            prefixes.database
        )
    }
}

fn connection_owner<'a>(nodes: &'a [Node], target: &'a Node) -> &'a Node {
    target
        .parent
        .map(|id| &nodes[id])
        .filter(|parent| parent.local == "database-description")
        .unwrap_or(target)
}

fn empty_element(prefix: &str, local: &str, attrs: Vec<(&str, String)>) -> String {
    let mut output = format!("<{prefix}:{local}");
    append_attrs(&mut output, prefix, attrs);
    output.push_str("/>");
    output
}

fn element_with_children(
    prefix: &str,
    local: &str,
    attrs: impl IntoIterator<Item = (&'static str, String)>,
    children: &str,
) -> String {
    let mut output = format!("<{prefix}:{local}");
    append_attrs(&mut output, prefix, attrs);
    if children.is_empty() {
        output.push_str("/>");
    } else {
        output.push('>');
        output.push_str(children);
        output.push_str(&format!("</{prefix}:{local}>"));
    }
    output
}

fn append_attrs<'a>(
    output: &mut String,
    prefix: &str,
    attrs: impl IntoIterator<Item = (&'a str, String)>,
) {
    for (name, value) in attrs {
        output.push(' ');
        output.push_str(prefix);
        output.push(':');
        output.push_str(name);
        output.push_str("=\"");
        output.push_str(&xml_escape(&value));
        output.push('"');
    }
}

fn push_bool<'a>(attrs: &mut Vec<(&'a str, String)>, name: &'a str, value: Option<bool>) {
    if let Some(value) = value {
        attrs.push((name, bool_text(value).to_owned()));
    }
}

fn push_option<'a, T: ToString>(
    attrs: &mut Vec<(&'a str, String)>,
    name: &'a str,
    value: Option<T>,
) {
    if let Some(value) = value {
        attrs.push((name, value.to_string()));
    }
}

fn push_action<'a>(
    attrs: &mut Vec<(&'a str, String)>,
    name: &'a str,
    value: Option<ReferentialAction>,
) {
    if let Some(value) = value {
        let value = match value {
            ReferentialAction::Cascade => "cascade",
            ReferentialAction::Restrict => "restrict",
            ReferentialAction::SetNull => "set-null",
            ReferentialAction::NoAction => "no-action",
            ReferentialAction::SetDefault => "set-default",
        };
        attrs.push((name, value.to_owned()));
    }
}

fn component_collection(kind: ComponentKind) -> &'static str {
    match kind {
        ComponentKind::Form => "forms",
        ComponentKind::Report => "reports",
    }
}

fn table_element(kind: TableKind) -> &'static str {
    match kind {
        TableKind::Representation => "table-representation",
        TableKind::Definition => "table-definition",
    }
}

fn column_element(table: &Node) -> &'static str {
    if table.local == "table-definition" {
        "column-definition"
    } else {
        "column"
    }
}

fn column_collection(table: &Node) -> &'static str {
    if table.local == "table-definition" {
        "column-definitions"
    } else {
        "columns"
    }
}

fn column_is_referenced(nodes: &[Node], table: &Node, name: &str) -> bool {
    descendants(nodes, table).any(|node| {
        matches!(node.local.as_str(), "key-column" | "index-column")
            && attribute(node, DATABASE_NAMESPACE, "name").is_some_and(|value| value == name)
    })
}

fn validate_extension(xml: &str) -> Result<()> {
    if xml.len() > MAX_VALUE_BYTES {
        return invalid("ODB producer extension exceeds the byte limit");
    }
    if xml.starts_with("<?xml") || xml.contains("<!DOCTYPE") {
        return invalid("ODB producer extension must be one inert element subtree");
    }
    litchi_odf_common::compact_xml::validate(xml.as_bytes()).map_err(Error::from)?;
    let nodes = scan(xml)?;
    let roots = nodes.iter().filter(|node| node.parent.is_none()).count();
    if roots != 1 {
        return invalid("ODB producer extension must contain exactly one element");
    }
    Ok(())
}

fn validate_name(value: &str, kind: &str) -> Result<()> {
    if value.is_empty() {
        return Err(Error::InvalidFormat(format!("ODB {kind} name is empty")));
    }
    validate_value(value, &format!("{kind} name"))
}

fn validate_value(value: &str, kind: &str) -> Result<()> {
    if value.len() > MAX_VALUE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODB {kind} exceeds the byte limit"
        )));
    }
    if value.chars().any(|character| {
        let scalar = u32::from(character);
        scalar == 0
            || scalar == 0xFFFE
            || scalar == 0xFFFF
            || (scalar < 0x20 && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(Error::InvalidFormat(format!(
            "ODB {kind} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}

fn xml_escape(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

const fn bool_text(value: bool) -> &'static str {
    if value { "true" } else { "false" }
}

fn str_from(value: &[u8]) -> &str {
    std::str::from_utf8(value).unwrap_or("")
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_owned()))
}
