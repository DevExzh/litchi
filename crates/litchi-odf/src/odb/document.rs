//! Namespace-aware, inert access to OpenDocument database front ends.

use crate::{OdfMetadata, OpenDocumentFamily, OpenDocumentPackage};
use litchi_core::{Error, Metadata, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 256;
const MAX_ATTRIBUTE_BYTES: usize = 1_048_576;
const MAX_TEXT_BYTES: usize = 32 * 1_048_576;

/// A recognized element in the standard ODF database vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DatabaseElementKind {
    Database,
    DataSource,
    ConnectionData,
    DatabaseDescription,
    FileBasedDatabase,
    ServerDatabase,
    ConnectionResource,
    Login,
    DriverSettings,
    ApplicationConnectionSettings,
    AutoIncrement,
    Delimiter,
    FontCharset,
    CharacterSet,
    TableSettings,
    TableSetting,
    TableFilter,
    TableIncludeFilter,
    TableExcludeFilter,
    TableFilterPattern,
    TableTypeFilter,
    TableType,
    DataSourceSettings,
    DataSourceSetting,
    DataSourceSettingValue,
    Forms,
    Reports,
    ComponentCollection,
    Component,
    Queries,
    QueryCollection,
    Query,
    OrderStatement,
    FilterStatement,
    UpdateTable,
    TableRepresentations,
    TableRepresentation,
    Columns,
    Column,
    SchemaDefinition,
    TableDefinitions,
    TableDefinition,
    ColumnDefinitions,
    ColumnDefinition,
    Keys,
    Key,
    KeyColumns,
    KeyColumn,
    Indices,
    Index,
    IndexColumns,
    IndexColumn,
    /// A future database element, embedded document, or vendor extension.
    Other,
}

/// One decoded attribute with its expanded namespace name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseAttribute {
    namespace_uri: Option<String>,
    local_name: String,
    value: String,
}

impl DatabaseAttribute {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn value(&self) -> &str {
        &self.value
    }
}

/// Ordered mixed content in the complete database subtree.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DatabaseContent {
    /// Decoded character content. Unknown declared entities remain in source notation.
    Text(String),
    /// A child element.
    Element(DatabaseElement),
}

/// One complete element from `office:database` and its descendants.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DatabaseElement {
    namespace_uri: Option<String>,
    local_name: String,
    attributes: Vec<DatabaseAttribute>,
    content: Vec<DatabaseContent>,
}

impl DatabaseElement {
    pub fn namespace_uri(&self) -> Option<&str> {
        self.namespace_uri.as_deref()
    }

    pub fn local_name(&self) -> &str {
        &self.local_name
    }

    pub fn kind(&self) -> DatabaseElementKind {
        if self.namespace_uri() == Some(OFFICE_NAMESPACE) && self.local_name == "database" {
            return DatabaseElementKind::Database;
        }
        if self.namespace_uri() != Some(DATABASE_NAMESPACE) {
            return DatabaseElementKind::Other;
        }
        match self.local_name.as_str() {
            "data-source" => DatabaseElementKind::DataSource,
            "connection-data" => DatabaseElementKind::ConnectionData,
            "database-description" => DatabaseElementKind::DatabaseDescription,
            "file-based-database" => DatabaseElementKind::FileBasedDatabase,
            "server-database" => DatabaseElementKind::ServerDatabase,
            "connection-resource" => DatabaseElementKind::ConnectionResource,
            "login" => DatabaseElementKind::Login,
            "driver-settings" => DatabaseElementKind::DriverSettings,
            "application-connection-settings" => DatabaseElementKind::ApplicationConnectionSettings,
            "auto-increment" => DatabaseElementKind::AutoIncrement,
            "delimiter" => DatabaseElementKind::Delimiter,
            "font-charset" => DatabaseElementKind::FontCharset,
            "character-set" => DatabaseElementKind::CharacterSet,
            "table-settings" => DatabaseElementKind::TableSettings,
            "table-setting" => DatabaseElementKind::TableSetting,
            "table-filter" => DatabaseElementKind::TableFilter,
            "table-include-filter" => DatabaseElementKind::TableIncludeFilter,
            "table-exclude-filter" => DatabaseElementKind::TableExcludeFilter,
            "table-filter-pattern" => DatabaseElementKind::TableFilterPattern,
            "table-type-filter" => DatabaseElementKind::TableTypeFilter,
            "table-type" => DatabaseElementKind::TableType,
            "data-source-settings" => DatabaseElementKind::DataSourceSettings,
            "data-source-setting" => DatabaseElementKind::DataSourceSetting,
            "data-source-setting-value" => DatabaseElementKind::DataSourceSettingValue,
            "forms" => DatabaseElementKind::Forms,
            "reports" => DatabaseElementKind::Reports,
            "component-collection" => DatabaseElementKind::ComponentCollection,
            "component" => DatabaseElementKind::Component,
            "queries" => DatabaseElementKind::Queries,
            "query-collection" => DatabaseElementKind::QueryCollection,
            "query" => DatabaseElementKind::Query,
            "order-statement" => DatabaseElementKind::OrderStatement,
            "filter-statement" => DatabaseElementKind::FilterStatement,
            "update-table" => DatabaseElementKind::UpdateTable,
            "table-representations" => DatabaseElementKind::TableRepresentations,
            "table-representation" => DatabaseElementKind::TableRepresentation,
            "columns" => DatabaseElementKind::Columns,
            "column" => DatabaseElementKind::Column,
            "schema-definition" => DatabaseElementKind::SchemaDefinition,
            "table-definitions" => DatabaseElementKind::TableDefinitions,
            "table-definition" => DatabaseElementKind::TableDefinition,
            "column-definitions" => DatabaseElementKind::ColumnDefinitions,
            "column-definition" => DatabaseElementKind::ColumnDefinition,
            "keys" => DatabaseElementKind::Keys,
            "key" => DatabaseElementKind::Key,
            "key-columns" => DatabaseElementKind::KeyColumns,
            "key-column" => DatabaseElementKind::KeyColumn,
            "indices" => DatabaseElementKind::Indices,
            "index" => DatabaseElementKind::Index,
            "index-columns" => DatabaseElementKind::IndexColumns,
            "index-column" => DatabaseElementKind::IndexColumn,
            _ => DatabaseElementKind::Other,
        }
    }

    pub fn attributes(&self) -> &[DatabaseAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name == local_name
            })
            .map(DatabaseAttribute::value)
    }

    pub fn content(&self) -> &[DatabaseContent] {
        &self.content
    }

    pub fn children(&self) -> impl Iterator<Item = &DatabaseElement> {
        self.content.iter().filter_map(|content| match content {
            DatabaseContent::Element(element) => Some(element),
            DatabaseContent::Text(_) => None,
        })
    }

    pub fn children_of_kind(
        &self,
        kind: DatabaseElementKind,
    ) -> impl Iterator<Item = &DatabaseElement> {
        self.children().filter(move |child| child.kind() == kind)
    }

    pub fn all_text(&self) -> String {
        fn append(element: &DatabaseElement, output: &mut String) {
            for content in &element.content {
                match content {
                    DatabaseContent::Text(text) => output.push_str(text),
                    DatabaseContent::Element(child) => append(child, output),
                }
            }
        }
        let mut output = String::new();
        append(self, &mut output);
        output
    }

    fn collect_kind<'a>(&'a self, kind: DatabaseElementKind, output: &mut Vec<&'a Self>) {
        if self.kind() == kind {
            output.push(self);
        }
        for child in self.children() {
            child.collect_kind(kind, output);
        }
    }
}

/// A validated OpenDocument Database Front End (`.odb`).
///
/// Connections, commands, queries, macros, embedded engines, forms, and reports
/// remain inert. This reader never connects to or executes a database.
pub struct DatabaseDocument {
    package: OpenDocumentPackage,
    database: DatabaseElement,
}

impl DatabaseDocument {
    /// Atomically replace or remove the packaged saved-query collection.
    ///
    /// Query commands, filters, URLs, embedded engines, and scripts remain
    /// inert. This method never connects, fetches, or executes package data.
    pub fn set_queries(
        &mut self,
        value: Option<&super::query::OdfDatabaseQueries>,
    ) -> Result<Option<super::query::OdfDatabaseQueries>> {
        let previous = self.queries()?;
        let content = self.package.content_xml()?;
        let content = super::query::set_database_queries_xml(&content, value)?;
        let package = self.package.with_replaced_content_xml(content)?;
        let staged = Self::from_bytes(package.into_bytes())?;
        staged.queries()?;
        *self = staged;
        Ok(previous)
    }

    /// Atomically replace or remove `db:schema-definition` in `content.xml`.
    pub fn set_schema_definition(
        &mut self,
        value: Option<&super::schema::OdfDatabaseSchemaDefinition>,
    ) -> Result<Option<super::schema::OdfDatabaseSchemaDefinition>> {
        if let Some(value) = value {
            value.validate()?;
        }
        let previous = self.schema_definition()?;
        let content = self.package.content_xml()?;
        let content = super::schema::set_database_schema_definition_xml(&content, value)?;
        let package = self.package.with_replaced_content_xml(content)?;
        let staged = Self::from_package(package)?;
        staged.schema_definition()?;
        *self = staged;
        Ok(previous)
    }

    pub fn clear_schema_definition(
        &mut self,
    ) -> Result<Option<super::schema::OdfDatabaseSchemaDefinition>> {
        self.set_schema_definition(None)
    }

    pub fn add_schema_table(
        &mut self,
        table: super::schema::OdfDatabaseTableDefinition,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let index = schema.tables.len();
            schema.tables.push(table);
            Ok(index)
        })
    }
    pub fn add_schema_view(
        &mut self,
        mut view: super::schema::OdfDatabaseTableDefinition,
    ) -> Result<usize> {
        view.table_type = Some("VIEW".to_string());
        self.add_schema_table(view)
    }
    pub fn update_schema_table(
        &mut self,
        index: usize,
        table: super::schema::OdfDatabaseTableDefinition,
    ) -> Result<super::schema::OdfDatabaseTableDefinition> {
        self.mutate_schema(|schema| replace_at(&mut schema.tables, index, table, "schema table"))
    }
    pub fn remove_schema_table(
        &mut self,
        index: usize,
    ) -> Result<super::schema::OdfDatabaseTableDefinition> {
        self.mutate_schema(|schema| remove_at(&mut schema.tables, index, "schema table"))
    }
    pub fn move_schema_table(&mut self, from: usize, to: usize) -> Result<()> {
        self.mutate_schema(|schema| move_at(&mut schema.tables, from, to, "schema table"))
    }

    pub fn add_schema_column(
        &mut self,
        table: usize,
        column: super::schema::OdfDatabaseColumnDefinition,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let values = &mut table_at(schema, table)?.columns;
            let index = values.len();
            values.push(column);
            Ok(index)
        })
    }
    pub fn update_schema_column(
        &mut self,
        table: usize,
        column: usize,
        value: super::schema::OdfDatabaseColumnDefinition,
    ) -> Result<super::schema::OdfDatabaseColumnDefinition> {
        self.mutate_schema(|schema| {
            replace_at(
                &mut table_at(schema, table)?.columns,
                column,
                value,
                "schema column",
            )
        })
    }
    pub fn remove_schema_column(
        &mut self,
        table: usize,
        column: usize,
    ) -> Result<super::schema::OdfDatabaseColumnDefinition> {
        self.mutate_schema(|schema| {
            remove_at(
                &mut table_at(schema, table)?.columns,
                column,
                "schema column",
            )
        })
    }
    pub fn move_schema_column(&mut self, table: usize, from: usize, to: usize) -> Result<()> {
        self.mutate_schema(|schema| {
            move_at(
                &mut table_at(schema, table)?.columns,
                from,
                to,
                "schema column",
            )
        })
    }

    pub fn add_schema_key(
        &mut self,
        table: usize,
        key: super::schema::OdfDatabaseKey,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let values = table_at(schema, table)?.keys.get_or_insert_with(Vec::new);
            let index = values.len();
            values.push(key);
            Ok(index)
        })
    }
    pub fn add_schema_relation(
        &mut self,
        table: usize,
        key: super::schema::OdfDatabaseKey,
    ) -> Result<usize> {
        if key.key_type != super::schema::OdfDatabaseKeyType::Foreign {
            return Err(Error::InvalidFormat(
                "schema relation must be a foreign key".into(),
            ));
        }
        self.add_schema_key(table, key)
    }
    pub fn update_schema_key(
        &mut self,
        table: usize,
        key: usize,
        value: super::schema::OdfDatabaseKey,
    ) -> Result<super::schema::OdfDatabaseKey> {
        self.mutate_schema(|schema| replace_at(keys_at(schema, table)?, key, value, "schema key"))
    }
    pub fn remove_schema_key(
        &mut self,
        table: usize,
        key: usize,
    ) -> Result<super::schema::OdfDatabaseKey> {
        self.mutate_schema(|schema| {
            let (removed, empty) = {
                let values = keys_at(schema, table)?;
                let removed = remove_at(values, key, "schema key")?;
                (removed, values.is_empty())
            };
            if empty {
                table_at(schema, table)?.keys = None;
            }
            Ok(removed)
        })
    }
    pub fn move_schema_key(&mut self, table: usize, from: usize, to: usize) -> Result<()> {
        self.mutate_schema(|schema| move_at(keys_at(schema, table)?, from, to, "schema key"))
    }
    pub fn add_schema_key_column(
        &mut self,
        table: usize,
        key: usize,
        group: usize,
        column: super::schema::OdfDatabaseKeyColumn,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let values = key_group_at(schema, table, key, group)?;
            let index = values.len();
            values.push(column);
            Ok(index)
        })
    }
    pub fn update_schema_key_column(
        &mut self,
        table: usize,
        key: usize,
        group: usize,
        column: usize,
        value: super::schema::OdfDatabaseKeyColumn,
    ) -> Result<super::schema::OdfDatabaseKeyColumn> {
        self.mutate_schema(|schema| {
            replace_at(
                key_group_at(schema, table, key, group)?,
                column,
                value,
                "schema key column",
            )
        })
    }
    pub fn remove_schema_key_column(
        &mut self,
        table: usize,
        key: usize,
        group: usize,
        column: usize,
    ) -> Result<super::schema::OdfDatabaseKeyColumn> {
        self.mutate_schema(|schema| {
            remove_at(
                key_group_at(schema, table, key, group)?,
                column,
                "schema key column",
            )
        })
    }
    pub fn move_schema_key_column(
        &mut self,
        table: usize,
        key: usize,
        group: usize,
        from: usize,
        to: usize,
    ) -> Result<()> {
        self.mutate_schema(|schema| {
            move_at(
                key_group_at(schema, table, key, group)?,
                from,
                to,
                "schema key column",
            )
        })
    }

    pub fn add_schema_index(
        &mut self,
        table: usize,
        index: super::schema::OdfDatabaseIndex,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let values = table_at(schema, table)?
                .indices
                .get_or_insert_with(Vec::new);
            let position = values.len();
            values.push(index);
            Ok(position)
        })
    }
    pub fn update_schema_index(
        &mut self,
        table: usize,
        index: usize,
        value: super::schema::OdfDatabaseIndex,
    ) -> Result<super::schema::OdfDatabaseIndex> {
        self.mutate_schema(|schema| {
            replace_at(indices_at(schema, table)?, index, value, "schema index")
        })
    }
    pub fn remove_schema_index(
        &mut self,
        table: usize,
        index: usize,
    ) -> Result<super::schema::OdfDatabaseIndex> {
        self.mutate_schema(|schema| {
            let (removed, empty) = {
                let values = indices_at(schema, table)?;
                let removed = remove_at(values, index, "schema index")?;
                (removed, values.is_empty())
            };
            if empty {
                table_at(schema, table)?.indices = None;
            }
            Ok(removed)
        })
    }
    pub fn move_schema_index(&mut self, table: usize, from: usize, to: usize) -> Result<()> {
        self.mutate_schema(|schema| move_at(indices_at(schema, table)?, from, to, "schema index"))
    }
    pub fn add_schema_index_column(
        &mut self,
        table: usize,
        index: usize,
        group: usize,
        column: super::schema::OdfDatabaseIndexColumn,
    ) -> Result<usize> {
        self.mutate_schema(|schema| {
            let values = index_group_at(schema, table, index, group)?;
            let position = values.len();
            values.push(column);
            Ok(position)
        })
    }
    pub fn update_schema_index_column(
        &mut self,
        table: usize,
        index: usize,
        group: usize,
        column: usize,
        value: super::schema::OdfDatabaseIndexColumn,
    ) -> Result<super::schema::OdfDatabaseIndexColumn> {
        self.mutate_schema(|schema| {
            replace_at(
                index_group_at(schema, table, index, group)?,
                column,
                value,
                "schema index column",
            )
        })
    }
    pub fn remove_schema_index_column(
        &mut self,
        table: usize,
        index: usize,
        group: usize,
        column: usize,
    ) -> Result<super::schema::OdfDatabaseIndexColumn> {
        self.mutate_schema(|schema| {
            remove_at(
                index_group_at(schema, table, index, group)?,
                column,
                "schema index column",
            )
        })
    }
    pub fn move_schema_index_column(
        &mut self,
        table: usize,
        index: usize,
        group: usize,
        from: usize,
        to: usize,
    ) -> Result<()> {
        self.mutate_schema(|schema| {
            move_at(
                index_group_at(schema, table, index, group)?,
                from,
                to,
                "schema index column",
            )
        })
    }

    fn mutate_schema<T>(
        &mut self,
        operation: impl FnOnce(&mut super::schema::OdfDatabaseSchemaDefinition) -> Result<T>,
    ) -> Result<T> {
        let mut schema = self.schema_definition()?.unwrap_or_default();
        let output = operation(&mut schema)?;
        schema.validate()?;
        self.set_schema_definition(Some(&schema))?;
        Ok(output)
    }

    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_package(OpenDocumentPackage::from_bytes(bytes)?)
    }

    pub fn from_bytes_with_password(bytes: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        Self::from_package(OpenDocumentPackage::from_bytes_with_password(
            bytes, password,
        )?)
    }

    pub fn open_with_password(path: impl AsRef<Path>, password: impl Into<String>) -> Result<Self> {
        Self::from_bytes_with_password(std::fs::read(path)?, password)
    }

    fn from_package(package: OpenDocumentPackage) -> Result<Self> {
        if package.family() != OpenDocumentFamily::Database {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument database: MIME type is '{}'",
                package.mimetype()
            )));
        }
        let database = parse_database_content(&package.content_xml()?)?;
        validate_database_root(&database)?;
        Ok(Self { package, database })
    }

    pub fn mimetype(&self) -> &str {
        self.package.mimetype()
    }

    pub fn database(&self) -> &DatabaseElement {
        &self.database
    }

    pub fn elements_of_kind(&self, kind: DatabaseElementKind) -> Vec<&DatabaseElement> {
        let mut elements = Vec::new();
        self.database.collect_kind(kind, &mut elements);
        elements
    }

    /// List opaque embedded database engine resources without interpreting them.
    pub fn embedded_database_files(&self) -> Result<Vec<String>> {
        let mut files: Vec<_> = self
            .package
            .files()?
            .into_iter()
            .filter(|path| path.starts_with("database/") && !path.ends_with('/'))
            .collect();
        files.sort_unstable();
        Ok(files)
    }

    /// Extract one opaque package part, including an embedded database resource.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.package.get_file(path)
    }

    pub fn metadata(&self) -> Result<Metadata> {
        self.package.metadata()
    }

    pub fn odf_metadata(&self) -> Result<Option<OdfMetadata>> {
        self.package.odf_metadata()
    }

    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.to_bytes()
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}

fn table_at(
    schema: &mut super::schema::OdfDatabaseSchemaDefinition,
    index: usize,
) -> Result<&mut super::schema::OdfDatabaseTableDefinition> {
    schema
        .tables
        .get_mut(index)
        .ok_or_else(|| Error::InvalidFormat("schema table index is out of bounds".into()))
}
fn keys_at(
    schema: &mut super::schema::OdfDatabaseSchemaDefinition,
    table: usize,
) -> Result<&mut Vec<super::schema::OdfDatabaseKey>> {
    table_at(schema, table)?
        .keys
        .as_mut()
        .ok_or_else(|| Error::InvalidFormat("schema table has no keys".into()))
}
fn indices_at(
    schema: &mut super::schema::OdfDatabaseSchemaDefinition,
    table: usize,
) -> Result<&mut Vec<super::schema::OdfDatabaseIndex>> {
    table_at(schema, table)?
        .indices
        .as_mut()
        .ok_or_else(|| Error::InvalidFormat("schema table has no indices".into()))
}
fn key_group_at(
    schema: &mut super::schema::OdfDatabaseSchemaDefinition,
    table: usize,
    key: usize,
    group: usize,
) -> Result<&mut Vec<super::schema::OdfDatabaseKeyColumn>> {
    keys_at(schema, table)?
        .get_mut(key)
        .ok_or_else(|| Error::InvalidFormat("schema key index is out of bounds".into()))?
        .column_groups
        .get_mut(group)
        .ok_or_else(|| {
            Error::InvalidFormat("schema key-column group index is out of bounds".into())
        })
}
fn index_group_at(
    schema: &mut super::schema::OdfDatabaseSchemaDefinition,
    table: usize,
    index: usize,
    group: usize,
) -> Result<&mut Vec<super::schema::OdfDatabaseIndexColumn>> {
    indices_at(schema, table)?
        .get_mut(index)
        .ok_or_else(|| Error::InvalidFormat("schema index is out of bounds".into()))?
        .column_groups
        .get_mut(group)
        .ok_or_else(|| {
            Error::InvalidFormat("schema index-column group index is out of bounds".into())
        })
}
fn replace_at<T>(values: &mut [T], index: usize, value: T, label: &str) -> Result<T> {
    let slot = values
        .get_mut(index)
        .ok_or_else(|| Error::InvalidFormat(format!("{label} index is out of bounds")))?;
    Ok(std::mem::replace(slot, value))
}
fn remove_at<T>(values: &mut Vec<T>, index: usize, label: &str) -> Result<T> {
    if index >= values.len() {
        return Err(Error::InvalidFormat(format!(
            "{label} index is out of bounds"
        )));
    }
    Ok(values.remove(index))
}
fn move_at<T>(values: &mut Vec<T>, from: usize, to: usize, label: &str) -> Result<()> {
    if from >= values.len() || to >= values.len() {
        return Err(Error::InvalidFormat(format!(
            "{label} reorder index is out of bounds"
        )));
    }
    if from != to {
        let value = values.remove(from);
        values.insert(to, value);
    }
    Ok(())
}

pub(super) fn parse_database_content(xml: &str) -> Result<DatabaseElement> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut body_depth = None;
    let mut database_seen = false;
    let mut database_depth = None;
    let mut complete = None;
    let mut stack = Vec::new();
    let mut node_count = 0usize;
    let mut text_bytes = 0usize;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid database XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || namespace_uri.as_deref() != Some(OFFICE_NAMESPACE)
                        || local != "document-content"
                    {
                        return Err(Error::InvalidFormat(
                            "database content must have one office:document-content root"
                                .to_string(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1
                    && namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                {
                    if body_seen {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                    body_depth = Some(2);
                } else if depth == 2 && body_depth == Some(2) {
                    if namespace_uri.as_deref() != Some(OFFICE_NAMESPACE)
                        || local != "database"
                        || database_seen
                    {
                        return Err(Error::InvalidFormat(
                            "database body must contain exactly one office:database".to_string(),
                        ));
                    }
                    database_seen = true;
                    database_depth = Some(3);
                }
                if database_depth.is_some() {
                    stack.push(make_element(
                        &reader,
                        namespace_uri,
                        local,
                        element,
                        &mut node_count,
                    )?);
                    if stack.len() > MAX_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "database nesting exceeds {MAX_DEPTH} levels"
                        )));
                    }
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("database XML nesting overflow".to_string())
                })?;
            },
            Event::Empty(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "database content root cannot be empty".to_string(),
                    ));
                }
                if depth == 1
                    && namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                {
                    if body_seen {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                } else if depth == 2 && body_depth == Some(2) {
                    return Err(Error::InvalidFormat(
                        "office:database cannot be empty".to_string(),
                    ));
                } else if database_depth.is_some() {
                    let node =
                        make_element(&reader, namespace_uri, local, element, &mut node_count)?;
                    stack
                        .last_mut()
                        .expect("database parent exists")
                        .content
                        .push(DatabaseContent::Element(node));
                }
            },
            Event::End(ref element) => {
                let namespace_uri = namespace_uri(&namespace)?;
                let local = decode_utf8(element.local_name().as_ref(), "element name")?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected database XML closing tag".to_string())
                })?;
                if database_depth.is_some() {
                    let node = stack.pop().ok_or_else(|| {
                        Error::InvalidFormat("database element stack underflow".to_string())
                    })?;
                    if stack.is_empty() {
                        complete = Some(node);
                        database_depth = None;
                    } else {
                        stack
                            .last_mut()
                            .expect("parent exists")
                            .content
                            .push(DatabaseContent::Element(node));
                    }
                }
                if namespace_uri.as_deref() == Some(OFFICE_NAMESPACE)
                    && local == "body"
                    && depth == 1
                {
                    body_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid database text: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::CData(ref text) if !stack.is_empty() => {
                let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                    Error::InvalidFormat(format!("invalid database CDATA: {error}"))
                })?;
                push_text(
                    stack.last_mut().expect("element exists"),
                    value.into_owned(),
                    &mut text_bytes,
                )?;
            },
            Event::GeneralRef(ref reference) if !stack.is_empty() => push_text(
                stack.last_mut().expect("element exists"),
                decode_reference(reference)?,
                &mut text_bytes,
            )?,
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the database root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the database root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || body_depth.is_some()
        || !database_seen
        || database_depth.is_some()
        || !stack.is_empty()
    {
        return Err(Error::InvalidFormat(
            "incomplete OpenDocument database structure".to_string(),
        ));
    }
    complete
        .ok_or_else(|| Error::InvalidFormat("database document has no office:database".to_string()))
}

fn make_element(
    reader: &NsReader<&[u8]>,
    namespace_uri: Option<String>,
    local_name: String,
    element: &BytesStart<'_>,
    node_count: &mut usize,
) -> Result<DatabaseElement> {
    *node_count = node_count
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat("database node count overflow".to_string()))?;
    if *node_count > MAX_NODES {
        return Err(Error::InvalidFormat(format!(
            "database exceeds {MAX_NODES} elements"
        )));
    }
    if element.attributes().count() > MAX_ATTRIBUTES {
        return Err(Error::InvalidFormat(format!(
            "database element exceeds {MAX_ATTRIBUTES} attributes"
        )));
    }
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid database attribute: {error}"))
        })?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let attribute_namespace = namespace_uri_from(&namespace)?;
        let attribute_local = decode_utf8(local.as_ref(), "attribute name")?;
        if attributes.iter().any(|existing: &DatabaseAttribute| {
            existing.namespace_uri == attribute_namespace && existing.local_name == attribute_local
        }) {
            return Err(Error::InvalidFormat(format!(
                "duplicate expanded database attribute '{attribute_local}'"
            )));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid database attribute value: {error}"))
            })?
            .into_owned();
        if value.len() > MAX_ATTRIBUTE_BYTES {
            return Err(Error::InvalidFormat(
                "database attribute exceeds 1 MiB".to_string(),
            ));
        }
        attributes.push(DatabaseAttribute {
            namespace_uri: attribute_namespace,
            local_name: attribute_local,
            value,
        });
    }
    Ok(DatabaseElement {
        namespace_uri,
        local_name,
        attributes,
        content: Vec::new(),
    })
}

pub(super) fn validate_database_root(database: &DatabaseElement) -> Result<()> {
    if database.kind() != DatabaseElementKind::Database {
        return Err(Error::InvalidFormat(
            "database subtree has the wrong root".to_string(),
        ));
    }
    let children: Vec<_> = database.children().collect();
    if children.first().map(|child| child.kind()) != Some(DatabaseElementKind::DataSource)
        || children
            .iter()
            .filter(|child| child.kind() == DatabaseElementKind::DataSource)
            .count()
            != 1
    {
        return Err(Error::InvalidFormat(
            "office:database must begin with exactly one db:data-source".to_string(),
        ));
    }
    Ok(())
}

fn push_text(element: &mut DatabaseElement, value: String, total: &mut usize) -> Result<()> {
    *total = total
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("database text size overflow".to_string()))?;
    if *total > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(
            "database exceeds 32 MiB of XML text".to_string(),
        ));
    }
    if let Some(DatabaseContent::Text(existing)) = element.content.last_mut() {
        existing.push_str(&value);
    } else {
        element.content.push(DatabaseContent::Text(value));
    }
    Ok(())
}

fn namespace_uri(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    namespace_uri_from(namespace)
}
fn namespace_uri_from(namespace: &ResolveResult<'_>) -> Result<Option<String>> {
    match namespace {
        ResolveResult::Unbound => Ok(None),
        ResolveResult::Bound(Namespace(uri)) => decode_utf8(uri, "namespace URI").map(Some),
        ResolveResult::Unknown(prefix) => Err(Error::InvalidFormat(format!(
            "unknown database namespace prefix '{}'",
            String::from_utf8_lossy(prefix)
        ))),
    }
}
fn decode_utf8(bytes: &[u8], kind: &str) -> Result<String> {
    std::str::from_utf8(bytes)
        .map(str::to_string)
        .map_err(|_| Error::InvalidFormat(format!("non-UTF-8 database {kind}")))
}
fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid database character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid database entity: {error}")))?;
    Ok(match name.as_ref() {
        "amp" => "&".to_string(),
        "lt" => "<".to_string(),
        "gt" => ">".to_string(),
        "quot" => "\"".to_string(),
        "apos" => "'".to_string(),
        _ => format!("&{name};"),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::constants;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str, embedded: bool) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        if embedded {
            writer.add_file("database/script", b"opaque SQL").unwrap();
            writer
                .add_file("database/properties", b"opaque config")
                .unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn database_xml() -> &'static str {
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"
 xmlns:x="http://www.w3.org/1999/xlink" xmlns:v="urn:vendor:db">
 <o:automatic-styles/><o:body><o:database>
  <d:data-source><d:connection-data><d:connection-resource x:href="sdbc:embedded:firebird" x:type="simple"/><d:login d:is-password-required="false"/></d:connection-data>
   <d:driver-settings d:parameter-name-substitution="false"/>
   <d:application-connection-settings d:max-row-count="100"><d:data-source-settings><d:data-source-setting d:data-source-setting-name="Type" d:data-source-setting-type="string"><d:data-source-setting-value>simple</d:data-source-setting-value></d:data-source-setting></d:data-source-settings></d:application-connection-settings>
  </d:data-source>
  <d:forms><d:component d:name="Entry" x:href="forms/Obj1"/></d:forms>
  <d:reports><d:component-collection d:name="Reports"><d:component d:name="Summary" x:href="reports/Obj1"/></d:component-collection></d:reports>
  <d:queries><d:query d:name="Recent" d:command="SELECT * FROM records" d:escape-processing="true"><d:order-statement d:command="id DESC"/></d:query></d:queries>
  <d:table-representations><d:table-representation d:name="records"><d:columns><d:column d:name="id"/><d:column d:name="title"/></d:columns></d:table-representation></d:table-representations>
  <v:engine v:mode="inert"><v:option>extension</v:option></v:engine>
 </o:database></o:body>
</o:document-content>"#
    }

    #[test]
    fn parses_libreoffice_style_database_subtree_and_opaque_engine_files() {
        let bytes = package(constants::ODF_DATABASE, database_xml(), true);
        let document = DatabaseDocument::from_bytes(bytes.clone()).unwrap();
        assert_eq!(document.database().kind(), DatabaseElementKind::Database);
        assert_eq!(
            document.elements_of_kind(DatabaseElementKind::ConnectionResource)[0]
                .attribute(Some("http://www.w3.org/1999/xlink"), "href"),
            Some("sdbc:embedded:firebird")
        );
        let queries = document.elements_of_kind(DatabaseElementKind::Query);
        assert_eq!(queries.len(), 1);
        assert_eq!(
            queries[0].attribute(Some(DATABASE_NAMESPACE), "command"),
            Some("SELECT * FROM records")
        );
        assert_eq!(
            document.elements_of_kind(DatabaseElementKind::DataSourceSettingValue)[0].all_text(),
            "simple"
        );
        assert_eq!(
            document
                .elements_of_kind(DatabaseElementKind::Component)
                .len(),
            2
        );
        assert_eq!(
            document.elements_of_kind(DatabaseElementKind::Other).len(),
            2
        );
        assert_eq!(
            document.embedded_database_files().unwrap(),
            ["database/properties", "database/script"]
        );
        assert_eq!(document.get_file("database/script").unwrap(), b"opaque SQL");
        assert_eq!(document.to_bytes(), bytes);
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn accepts_readers_and_file_based_sources_with_arbitrary_prefixes() {
        let xml = format!(
            r#"<x:document-content xmlns:x="{OFFICE_NAMESPACE}" xmlns:q="{DATABASE_NAMESPACE}" xmlns:l="http://www.w3.org/1999/xlink"><x:body><x:database><q:data-source><q:connection-data><q:database-description><q:file-based-database l:href="$(userurl)/database/biblio" q:media-type="application/dbase"/></q:database-description></q:connection-data></q:data-source></x:database></x:body></x:document-content>"#
        );
        let bytes = package(constants::ODF_DATABASE, &xml, false);
        let document = DatabaseDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        let file = &document.elements_of_kind(DatabaseElementKind::FileBasedDatabase)[0];
        assert_eq!(
            file.attribute(Some(DATABASE_NAMESPACE), "media-type"),
            Some("application/dbase")
        );
        assert_eq!(document.into_bytes(), bytes);
    }

    #[test]
    fn rejects_other_families_and_invalid_database_hierarchy() {
        assert!(
            DatabaseDocument::from_bytes(package(constants::ODF_TEXT, database_xml(), false))
                .is_err()
        );
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:queries/></o:database></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/><d:data-source/></o:database></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:text><d:data-source/></o:text></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0"><o:body><o:database><d:data-source/></o:database></o:body>"#,
        ] {
            assert!(
                DatabaseDocument::from_bytes(package(constants::ODF_DATABASE, xml, false)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_duplicate_expanded_attributes_and_excessive_nesting() {
        let duplicate = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DATABASE_NAMESPACE}" xmlns:a="urn:test" xmlns:b="urn:test"><o:body><o:database><d:data-source a:x="one" b:x="two"/></o:database></o:body></o:document-content>"#
        );
        assert!(
            DatabaseDocument::from_bytes(package(constants::ODF_DATABASE, &duplicate, false))
                .is_err()
        );
        let nested = "<v:x>".repeat(MAX_DEPTH) + &"</v:x>".repeat(MAX_DEPTH);
        let deep = format!(
            r#"<o:document-content xmlns:o="{OFFICE_NAMESPACE}" xmlns:d="{DATABASE_NAMESPACE}" xmlns:v="urn:vendor"><o:body><o:database><d:data-source>{nested}</d:data-source></o:database></o:body></o:document-content>"#
        );
        assert!(
            DatabaseDocument::from_bytes(package(constants::ODF_DATABASE, &deep, false)).is_err()
        );
    }
}
