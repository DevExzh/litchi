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
use litchi_odf_common::package::edit::Addition;
use quick_xml::{
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{collections::BTreeSet, ops::Range};

use crate::{
    ApplicationConnectionSettings, AutoIncrementSettings, CharacterSetSettings, Column, Component,
    ComponentKind, Connection, DataSourceSetting, Database, DatabaseSettings, DelimiterSettings,
    DriverSettings, FileDatabaseTarget, Index, Key, LoginSettings, Query, ReferentialAction,
    ServerDatabaseAddress, ServerDatabaseTarget, Table, TableFilter, TableKind, TableSetting,
    TableTypeFilter,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DATABASE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:database:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const MAX_VALUE_BYTES: usize = 1024 * 1024;
const MAX_OPERATIONS: usize = 65_536;
const MAX_TABLE_SETTINGS: usize = 65_536;
const MAX_OUTPUT_BYTES: usize = 256 * 1024 * 1024;
const MAX_SETTINGS_VALUE_COUNT: usize = 65_536;
const MAX_SETTINGS_OUTPUT_BYTES: usize = 64 * 1024 * 1024;
const MAX_XML_DEPTH: usize = 256;
const MAX_SCAN_NODES: usize = 1_000_000;
const MAX_SCAN_ATTRIBUTES: usize = 1_000_000;

/// The semantic family changed by one staged operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChangeKind {
    Connection,
    Settings,
    Query,
    Table,
    Column,
    Key,
    Index,
    Component,
    ProducerExtension,
}

/// The CRUD direction of one semantic effect.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChangeAction {
    Create,
    Update,
    Remove,
}

/// One ordered semantic effect in a unified transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Change {
    action: ChangeAction,
    kind: ChangeKind,
    target: String,
}

impl Change {
    /// Returns whether the target is created, updated, or removed.
    #[must_use]
    pub const fn action(&self) -> ChangeAction {
        self.action
    }

    /// Returns the semantic family affected by this operation.
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        self.kind
    }

    /// Returns the stable collision-free semantic target key.
    ///
    /// Nested owners use a decimal owner-byte-length prefix so producer names
    /// containing separators remain unambiguous.
    #[must_use]
    pub fn target(&self) -> &str {
        &self.target
    }
}

/// A source-bound unified edit over one immutable database snapshot.
///
/// All values remain inert XML metadata. No method opens a connection, loads a
/// driver, follows a component link, or executes a stored command.
#[derive(Clone)]
pub struct Edit<'source> {
    source: &'source Database,
    policy: crate::EditPolicy,
    content: String,
    changes: Vec<Change>,
    legacy_query: Option<QueryChange>,
    payload_additions: Vec<Addition>,
    payload_directories: Vec<(String, String)>,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Database, policy: crate::EditPolicy) -> Self {
        Self {
            source,
            policy,
            content: source.content_xml().to_owned(),
            changes: Vec::new(),
            legacy_query: None,
            payload_additions: Vec::new(),
            payload_directories: Vec::new(),
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

    /// Copies one bounded inert table declaration from another snapshot.
    ///
    /// Foreign-key targets remain semantic names; this operation never opens
    /// either database connection or copies database rows.
    ///
    /// # Errors
    ///
    /// Returns an error when the source catalog is invalid, the table is
    /// absent, or the destination cannot accept the declaration.
    pub fn transfer_table_from(&mut self, source: &Database, name: &str) -> Result<()> {
        self.transfer_table_from_with(source, name, crate::DependencyDisposition::Refuse)
    }

    /// Copies one table with an explicit foreign-table dependency policy.
    /// `Cascade` recursively copies missing referenced tables; `Refuse` never
    /// creates an unresolved relation.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing/duplicate table, invalid source schema,
    /// unresolved dependency, cycle that cannot be represented, or limit.
    pub fn transfer_table_from_with(
        &mut self,
        source: &Database,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?.to_owned();
        self.stage_atomically(|candidate| {
            if candidate.staged_table(name)?.is_some() {
                return invalid("ODB transfer table already exists in destination");
            }
            let mut visiting = BTreeSet::new();
            candidate.transfer_table_from_catalog(&catalog, name, disposition, &mut visiting)
        })
    }

    /// Copies one bounded inert stored-query declaration from another
    /// snapshot without parsing or executing its command.
    ///
    /// # Errors
    ///
    /// Returns an error when the source catalog is invalid, the query is
    /// absent, or the destination cannot accept the declaration.
    pub fn transfer_query_from(&mut self, source: &Database, name: &str) -> Result<()> {
        let catalog = source.catalog()?;
        let query = catalog
            .query(name)?
            .ok_or_else(|| Error::InvalidFormat("ODB transfer query does not exist".to_string()))?;
        let query = query.clone();
        self.stage_atomically(|candidate| candidate.add_query(query))
    }

    /// Copies one stored query with an explicit local-table dependency policy.
    ///
    /// Common `SELECT`, `INSERT`, `UPDATE`, and `DELETE` shapes use
    /// [`Query::command_semantics`] to discover exact relation names.
    /// `Cascade` copies missing donor tables and their foreign-key closure;
    /// `Refuse` requires every dependency to exist already. SQL shapes whose
    /// dependency graph is not conservatively complete are refused.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/duplicate query, unsupported SQL
    /// dependency shape, unresolved donor relation, destination collision, or
    /// table dependency failure. No command or database content is executed.
    pub fn transfer_query_from_with(
        &mut self,
        source: &Database,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?.to_owned();
        let query = owned_query(&catalog, name)?.clone();
        let semantics = query.command_semantics()?;
        if let crate::QueryDependencySupport::Unsupported(reason) = semantics.dependency_support() {
            return Err(Error::Unsupported(format!(
                "ODB query dependency closure is unsupported: {reason:?}"
            )));
        }
        let mut dependencies = semantics
            .relations()
            .iter()
            .map(|relation| relation.name().to_owned())
            .collect::<BTreeSet<_>>();
        if let Some(target) = query.update_target() {
            dependencies.insert(target.name().to_owned());
        }
        self.stage_atomically(|candidate| {
            if candidate.staged_catalog()?.query(name)?.is_some() {
                return invalid("ODB transfer query already exists in destination");
            }
            let mut visiting = BTreeSet::new();
            for dependency in dependencies {
                if candidate.staged_table(&dependency)?.is_some() {
                    continue;
                }
                if disposition == crate::DependencyDisposition::Refuse {
                    return dependency_refusal(
                        "ODB query transfer requires a local relation dependency",
                    );
                }
                if owned_table(&catalog, &dependency).is_err() {
                    return dependency_refusal(
                        "ODB query transfer cannot close an external or qualified relation",
                    );
                }
                candidate.transfer_table_from_catalog(
                    &catalog,
                    &dependency,
                    crate::DependencyDisposition::Cascade,
                    &mut visiting,
                )?;
            }
            candidate.add_query(query)
        })
    }

    /// Copies the inert connection declaration, including a file or resource
    /// IRI, without opening it. An absent declaration removes the destination
    /// declaration.
    ///
    /// # Errors
    ///
    /// Returns an error when the source catalog is invalid or the destination
    /// cannot accept the bounded declaration.
    pub fn transfer_connection_from(&mut self, source: &Database) -> Result<()> {
        let connection = source.catalog()?.connection().cloned();
        self.stage_atomically(|candidate| candidate.set_connection_inner(connection))
    }

    /// Copies one column from the same named source table.
    ///
    /// # Errors
    ///
    /// Returns an error when either table/column is absent, source metadata is
    /// invalid, or the destination cannot accept the declaration.
    pub fn transfer_column_from(
        &mut self,
        source: &Database,
        table: &str,
        name: &str,
    ) -> Result<()> {
        let catalog = source.catalog()?;
        let source_table = catalog
            .table(table)?
            .ok_or_else(|| Error::InvalidFormat("ODB transfer table does not exist".to_string()))?;
        let column = source_table
            .columns()
            .iter()
            .find(|column| column.name() == name)
            .ok_or_else(|| {
                Error::InvalidFormat("ODB transfer column does not exist".to_string())
            })?;
        let column = column.clone();
        self.stage_atomically(|candidate| candidate.add_column(table, column))
    }

    /// Copies one key/relation with explicit dependency closure.
    ///
    /// `Cascade` copies missing local columns and a missing referenced table;
    /// `Refuse` reports those dependencies without mutation.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent selector, invalid dependency, conflict,
    /// or exceeded bound.
    pub fn transfer_key_from(
        &mut self,
        source: &Database,
        table: &str,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?.to_owned();
        let source_table = owned_table(&catalog, table)?;
        let key = source_table
            .keys()
            .iter()
            .find(|key| key.name() == Some(name))
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("ODB transfer key does not exist".to_string()))?;
        self.stage_atomically(|candidate| {
            for column in key.columns() {
                let Some(column_name) = column.name() else {
                    return invalid("ODB transfer key has an unnamed local column");
                };
                if !candidate.staged_column_exists(table, column_name)? {
                    if disposition == crate::DependencyDisposition::Refuse {
                        return dependency_refusal("ODB transfer key requires a local column");
                    }
                    let source_column = source_table
                        .columns()
                        .iter()
                        .find(|column| column.name() == column_name)
                        .cloned()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODB transfer key local column does not exist".to_string(),
                            )
                        })?;
                    candidate.add_column(table, source_column)?;
                }
            }
            if let Some(referenced) = key.referenced_table()
                && candidate.staged_table(referenced)?.is_none()
            {
                if disposition == crate::DependencyDisposition::Refuse {
                    return dependency_refusal("ODB transfer key requires a referenced table");
                }
                let mut visiting = BTreeSet::new();
                candidate.transfer_table_from_catalog(
                    &catalog,
                    referenced,
                    crate::DependencyDisposition::Cascade,
                    &mut visiting,
                )?;
            }
            candidate.validate_key_against_staged(table, &key)?;
            candidate.add_key(table, key)
        })
    }

    /// Copies one index and optionally closes missing local-column dependencies.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent selector, invalid dependency, conflict,
    /// or exceeded bound.
    pub fn transfer_index_from(
        &mut self,
        source: &Database,
        table: &str,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?.to_owned();
        let source_table = owned_table(&catalog, table)?;
        let index = source_table
            .indices()
            .iter()
            .find(|index| index.name() == name)
            .cloned()
            .ok_or_else(|| Error::InvalidFormat("ODB transfer index does not exist".to_string()))?;
        self.stage_atomically(|candidate| {
            for column in index.columns() {
                if !candidate.staged_column_exists(table, column.name())? {
                    if disposition == crate::DependencyDisposition::Refuse {
                        return dependency_refusal("ODB transfer index requires a local column");
                    }
                    let source_column = source_table
                        .columns()
                        .iter()
                        .find(|candidate| candidate.name() == column.name())
                        .cloned()
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "ODB transfer index local column does not exist".to_string(),
                            )
                        })?;
                    candidate.add_column(table, source_column)?;
                }
            }
            candidate.validate_index_against_staged(table, &index)?;
            candidate.add_index(table, index)
        })
    }

    /// Copies one unambiguous inert form/report declaration from another
    /// snapshot. A local linked package subtree is copied byte-for-byte to the
    /// same path; an external IRI remains inert and is never followed.
    ///
    /// # Errors
    ///
    /// Returns an error when the source catalog is invalid, the selector is
    /// absent or ambiguous, or the destination cannot accept the declaration.
    pub fn transfer_component_from(
        &mut self,
        source: &Database,
        kind: ComponentKind,
        name: &str,
    ) -> Result<()> {
        self.transfer_component_from_with(
            source,
            kind,
            name,
            crate::ActiveContentDisposition::Refuse,
        )
    }

    /// Copies one component to the same package path with explicit handling
    /// for active-content dependencies in its local linked subtree.
    pub fn transfer_component_from_with(
        &mut self,
        source: &Database,
        kind: ComponentKind,
        name: &str,
        active_content: crate::ActiveContentDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODB transfer component does not exist".to_string())
        })?;
        if matches.next().is_some() {
            return invalid("ODB transfer component selector is ambiguous");
        }
        let component = component.clone();
        self.stage_atomically(|candidate| {
            candidate.stage_component_payload(source, &component, None, active_content)?;
            candidate.add_component(component)
        })
    }

    /// Copies one component and explicitly remaps its local linked package
    /// subtree to `destination_href`.
    ///
    /// The destination must be a relative package path and must not collide
    /// with an existing or already-staged member. External IRIs are never
    /// dereferenced.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous component, an external source
    /// IRI, an unsafe destination, a missing payload, or a package collision.
    pub fn transfer_component_from_to(
        &mut self,
        source: &Database,
        kind: ComponentKind,
        name: &str,
        destination_href: &str,
    ) -> Result<()> {
        self.transfer_component_from_to_with(
            source,
            kind,
            name,
            destination_href,
            crate::ActiveContentDisposition::Refuse,
        )
    }

    /// Copies and remaps one component with explicit handling for active
    /// declarations contained in its local linked package subtree.
    pub fn transfer_component_from_to_with(
        &mut self,
        source: &Database,
        kind: ComponentKind,
        name: &str,
        destination_href: &str,
        active_content: crate::ActiveContentDisposition,
    ) -> Result<()> {
        let catalog = source.catalog()?;
        let mut matches = catalog
            .components()
            .iter()
            .filter(|component| component.kind() == kind && component.name() == Some(name));
        let component = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODB transfer component does not exist".to_string())
        })?;
        if matches.next().is_some() {
            return invalid("ODB transfer component selector is ambiguous");
        }
        let component = component.clone();
        self.stage_atomically(|candidate| {
            candidate.stage_component_payload(
                source,
                &component,
                Some(destination_href),
                active_content,
            )?;
            candidate.add_component(component.with_href(destination_href))
        })
    }

    fn stage_component_payload(
        &mut self,
        source: &Database,
        component: &Component,
        destination_href: Option<&str>,
        active_content: crate::ActiveContentDisposition,
    ) -> Result<()> {
        if source.protection_status()?.is_encrypted() {
            return Err(Error::Unsupported(
                "ODB component payload transfer from an encrypted donor is refused".to_string(),
            ));
        }
        let Some(source_href) = component.href() else {
            if destination_href.is_some() {
                return invalid("ODB component has no linked package payload to remap");
            }
            return Ok(());
        };
        let Some(source_prefix) = local_component_prefix(source_href)? else {
            if destination_href.is_some() {
                return invalid("ODB external component IRI cannot be remapped or followed");
            }
            return Ok(());
        };
        let destination = destination_href.unwrap_or(source_href);
        let destination_prefix = local_component_prefix(destination)?.ok_or_else(|| {
            Error::InvalidFormat(
                "ODB component payload destination must be a relative package path".to_string(),
            )
        })?;
        if self
            .source
            .package
            .contains_package_prefix(&destination_prefix)?
            || self
                .payload_additions
                .iter()
                .any(|addition| addition.path.starts_with(&destination_prefix))
            || self
                .payload_directories
                .iter()
                .any(|(path, _)| path.starts_with(&destination_prefix))
        {
            return invalid("ODB component payload destination already exists");
        }
        let (additions, directories) = source
            .package
            .component_payload(&source_prefix, &destination_prefix)?;
        if additions.is_empty() && directories.is_empty() {
            return invalid("ODB linked component package payload does not exist");
        }
        if let crate::ComponentTransferSupport::Refused(reason) =
            crate::package::Snapshot::payload_transfer_support(&additions)
        {
            return Err(Error::Unsupported(format!(
                "ODB component payload transfer is refused: {reason:?}"
            )));
        }
        let active_dependencies =
            crate::package::Snapshot::payload_active_content_count(&additions)?;
        if active_dependencies > 0 && active_content == crate::ActiveContentDisposition::Refuse {
            return Err(Error::Unsupported(format!(
                "ODB component payload has {active_dependencies} active-content dependencies"
            )));
        }
        for addition in additions {
            if let Some(existing) = self
                .payload_additions
                .iter()
                .find(|existing| existing.path == addition.path)
            {
                if existing.bytes != addition.bytes || existing.media_type != addition.media_type {
                    return invalid("ODB component package dependencies conflict");
                }
                continue;
            }
            match self.source.package.file_matches(&addition)? {
                Some(true) => {},
                Some(false) => {
                    return invalid("ODB component package dependency collides in destination");
                },
                None => self.payload_additions.push(addition),
            }
        }
        for (path, media_type) in directories {
            if let Some((_, existing)) = self
                .payload_directories
                .iter()
                .find(|(existing, _)| existing == &path)
            {
                if existing != &media_type {
                    return invalid("ODB component package directory dependencies conflict");
                }
                continue;
            }
            match self.source.package.directory_media_type(&path)? {
                Some(existing) if existing == media_type => {},
                Some(_) => {
                    return invalid("ODB component package directory collides in destination");
                },
                None => self.payload_directories.push((path, media_type)),
            }
        }
        Ok(())
    }

    /// Copies one exact, compact, inert producer-extension subtree.
    ///
    /// # Errors
    ///
    /// Returns an error when the source selector is absent/ambiguous or the
    /// destination rejects the compact authored fragment.
    pub fn transfer_producer_extension_from(
        &mut self,
        source: &Database,
        namespace: &str,
        local: &str,
    ) -> Result<()> {
        let extensions = source.producer_extensions()?;
        let mut matches = extensions.iter().filter(|extension| {
            extension.namespace() == namespace && extension.local_name() == local
        });
        let extension = matches.next().ok_or_else(|| {
            Error::InvalidFormat("ODB transfer producer extension does not exist".to_string())
        })?;
        if matches.next().is_some() {
            return invalid("ODB transfer producer extension selector is ambiguous");
        }
        let xml = extension.xml().to_owned();
        self.stage_atomically(|candidate| candidate.add_producer_extension(&xml))
    }

    fn stage_atomically(&mut self, operation: impl FnOnce(&mut Self) -> Result<()>) -> Result<()> {
        let mut candidate = self.clone();
        operation(&mut candidate)?;
        *self = candidate;
        Ok(())
    }

    fn staged_table(&self, name: &str) -> Result<Option<Table>> {
        self.staged_catalog()?
            .table(name)
            .map(<Option<&Table>>::cloned)
    }

    fn staged_column_exists(&self, table: &str, name: &str) -> Result<bool> {
        self.staged_catalog()?.table(table)?.map_or_else(
            || {
                Err(Error::InvalidFormat(
                    "ODB destination table does not exist".to_string(),
                ))
            },
            |table| Ok(table.columns().iter().any(|column| column.name() == name)),
        )
    }

    fn staged_catalog(&self) -> Result<crate::Catalog<'_>> {
        crate::Catalog::parse(&self.content, crate::Limits::default())
    }

    fn transfer_table_from_catalog(
        &mut self,
        catalog: &crate::OwnedCatalog,
        name: &str,
        disposition: crate::DependencyDisposition,
        visiting: &mut BTreeSet<String>,
    ) -> Result<()> {
        if self.staged_table(name)?.is_some() || !visiting.insert(name.to_owned()) {
            return Ok(());
        }
        let table = owned_table(catalog, name)?.clone();
        validate_source_table_dependencies(catalog, &table)?;
        for key in table.keys() {
            let Some(referenced) = key.referenced_table() else {
                continue;
            };
            if referenced == name || self.staged_table(referenced)?.is_some() {
                continue;
            }
            if disposition == crate::DependencyDisposition::Refuse {
                return dependency_refusal("ODB table transfer requires a referenced table");
            }
            self.transfer_table_from_catalog(catalog, referenced, disposition, visiting)?;
        }
        self.add_table(table)?;
        visiting.remove(name);
        Ok(())
    }

    fn validate_key_against_staged(&self, table: &str, key: &Key) -> Result<()> {
        let catalog = self.staged_catalog()?;
        let owner = catalog.table(table)?.ok_or_else(|| {
            Error::InvalidFormat("ODB destination table does not exist".to_string())
        })?;
        validate_key_columns(owner, key)?;
        if let Some(referenced) = key.referenced_table() {
            let target = catalog.table(referenced)?.ok_or_else(|| {
                Error::InvalidFormat("ODB referenced destination table does not exist".to_string())
            })?;
            validate_related_columns(target, key)?;
        }
        Ok(())
    }

    fn validate_index_against_staged(&self, table: &str, index: &Index) -> Result<()> {
        let catalog = self.staged_catalog()?;
        let owner = catalog.table(table)?.ok_or_else(|| {
            Error::InvalidFormat("ODB destination table does not exist".to_string())
        })?;
        if index
            .columns()
            .iter()
            .any(|index_column| !has_column(owner, index_column.name()))
        {
            return invalid("ODB index references a missing local column");
        }
        Ok(())
    }

    /// Adds a table presentation or schema definition at the collection tail.
    pub fn add_table(&mut self, table: Table) -> Result<()> {
        validate_table(&table)?;
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
        self.record(ChangeAction::Create, ChangeKind::Table, table.name())
    }

    /// Replaces one complete table declaration while preserving adjacent XML.
    pub fn replace_table(&mut self, name: &str, table: Table) -> Result<()> {
        validate_table(&table)?;
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
        self.record(ChangeAction::Update, ChangeKind::Table, name)
    }

    /// Removes one unambiguous table declaration.
    pub fn remove_table(&mut self, name: &str) -> Result<()> {
        self.remove_table_with(name, crate::DependencyDisposition::Refuse)
    }

    /// Removes a table under an explicit modeled-dependency disposition.
    pub fn remove_table_with(
        &mut self,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_named_table(&nodes, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        let incoming = nodes
            .iter()
            .filter(|node| {
                node.local == "key"
                    && attribute(node, DATABASE_NAMESPACE, "referenced-table-name")
                        .is_some_and(|value| value == name)
                    && !ancestors(&nodes, node).any(|ancestor| ancestor.id == site.id)
            })
            .collect::<Vec<_>>();
        if !incoming.is_empty() && disposition == crate::DependencyDisposition::Refuse {
            return Err(Error::Unsupported(
                "ODB table removal requires an incoming-relation disposition".to_string(),
            ));
        }
        if disposition == crate::DependencyDisposition::Cascade {
            let mut edits = incoming
                .iter()
                .map(|key| TextEdit {
                    range: key.full.clone(),
                    value: String::new(),
                })
                .collect::<Vec<_>>();
            edits.push(TextEdit {
                range: site.full.clone(),
                value: String::new(),
            });
            apply_edits(&mut self.content, edits)?;
            for key in incoming {
                let owner = ancestors(&nodes, key)
                    .find(|node| node.local == "table-definition")
                    .and_then(|node| attribute(node, DATABASE_NAMESPACE, "name"));
                let key_name = attribute(key, DATABASE_NAMESPACE, "name");
                if let (Some(owner), Some(key_name)) = (owner, key_name) {
                    self.record(
                        ChangeAction::Remove,
                        ChangeKind::Key,
                        &child_target(owner, key_name),
                    )?;
                }
            }
        } else {
            replace_span(&mut self.content, site.full.clone(), "")?;
        }
        self.record(ChangeAction::Remove, ChangeKind::Table, name)
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
        self.record(ChangeAction::Update, ChangeKind::Table, name)
    }

    /// Adds a column to a table at the collection tail.
    pub fn add_column(&mut self, table: &str, column: Column) -> Result<()> {
        validate_column(&column)?;
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
        self.record(
            ChangeAction::Create,
            ChangeKind::Column,
            &child_target(table, column.name()),
        )
    }

    /// Replaces one complete column declaration.
    pub fn replace_column(&mut self, table: &str, name: &str, column: Column) -> Result<()> {
        validate_column(&column)?;
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
        self.record(
            ChangeAction::Update,
            ChangeKind::Column,
            &child_target(table, name),
        )
    }

    /// Removes one table column.
    pub fn remove_column(&mut self, table: &str, name: &str) -> Result<()> {
        self.remove_column_with(table, name, crate::DependencyDisposition::Refuse)
    }

    /// Removes a column under an explicit key/index dependency disposition.
    pub fn remove_column_with(
        &mut self,
        table: &str,
        name: &str,
        disposition: crate::DependencyDisposition,
    ) -> Result<()> {
        let nodes = scan(&self.content)?;
        let table_site = find_named_table(&nodes, table)?
            .ok_or_else(|| Error::InvalidFormat("ODB table does not exist".to_string()))?;
        let site = find_named_child(&nodes, table_site, column_element(table_site), name)?
            .ok_or_else(|| Error::InvalidFormat("ODB column does not exist".to_string()))?;
        let dependents = column_dependents(&nodes, table_site, table, name);
        if !dependents.is_empty() && disposition == crate::DependencyDisposition::Refuse {
            return Err(Error::Unsupported(
                "ODB column removal requires a key/index dependency disposition".to_string(),
            ));
        }
        if disposition == crate::DependencyDisposition::Cascade {
            let mut edits = dependents
                .iter()
                .map(|dependent| TextEdit {
                    range: dependent.node.full.clone(),
                    value: String::new(),
                })
                .collect::<Vec<_>>();
            edits.push(TextEdit {
                range: site.full.clone(),
                value: String::new(),
            });
            apply_edits(&mut self.content, edits)?;
            for dependent in dependents {
                self.record(
                    ChangeAction::Remove,
                    dependent.kind,
                    &child_target(dependent.table, dependent.name),
                )?;
            }
        } else {
            replace_span(&mut self.content, site.full.clone(), "")?;
        }
        self.record(
            ChangeAction::Remove,
            ChangeKind::Column,
            &child_target(table, name),
        )
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
        self.record(
            ChangeAction::Update,
            ChangeKind::Column,
            &child_target(table, name),
        )
    }

    /// Adds a schema key. Foreign keys are the ODF relation representation.
    pub fn add_key(&mut self, table: &str, key: Key) -> Result<()> {
        let name = key
            .name()
            .ok_or_else(|| Error::InvalidFormat("ODB authored key requires a name".to_string()))?;
        validate_key(&key)?;
        self.validate_key_against_staged(table, &key)?;
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
        validate_key(&key)?;
        self.validate_key_against_staged(table, &key)?;
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
        validate_index(&index)?;
        self.validate_index_against_staged(table, &index)?;
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
        validate_index(&index)?;
        self.validate_index_against_staged(table, &index)?;
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
        validate_query(&query)?;
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
        self.record(ChangeAction::Create, ChangeKind::Query, query.name())
    }

    /// Replaces a stored query without interpreting its command.
    pub fn replace_query(&mut self, name: &str, query: Query) -> Result<()> {
        validate_query(&query)?;
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
        self.record(ChangeAction::Update, ChangeKind::Query, name)
    }

    /// Removes a stored query.
    pub fn remove_query(&mut self, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_named_node(&nodes, "query", name)?
            .ok_or_else(|| Error::InvalidFormat("ODB query does not exist".to_string()))?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(ChangeAction::Remove, ChangeKind::Query, name)
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
        self.record(ChangeAction::Update, ChangeKind::Query, name)
    }

    /// Sets or removes the stored `db:escape-processing` declaration.
    pub fn set_query_escape_processing(&mut self, name: &str, value: Option<bool>) -> Result<()> {
        self.prepare_legacy_query(name)?;
        let lexical = value.map(bool_text);
        self.set_named_db_attribute("query", name, "escape-processing", lexical)?;
        if let Some(change) = self.legacy_query.as_mut() {
            change.after_escape_processing = value;
        }
        self.record(ChangeAction::Update, ChangeKind::Query, name)
    }

    /// Replaces, creates, or removes the inert database connection target.
    pub fn set_connection(&mut self, connection: Option<Connection>) -> Result<()> {
        self.stage_atomically(|candidate| candidate.set_connection_inner(connection))
    }

    fn set_connection_inner(&mut self, connection: Option<Connection>) -> Result<()> {
        if let Some(value) = connection.as_ref() {
            validate_connection(value)?;
        }
        if self
            .staged_catalog()
            .ok()
            .and_then(|catalog| catalog.connection().cloned())
            == connection
        {
            return Ok(());
        }
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
        let action = match (targets.is_empty(), connection.is_some()) {
            (true, true) => ChangeAction::Create,
            (false, true) => ChangeAction::Update,
            (false, false) => ChangeAction::Remove,
            (true, false) => return Ok(()),
        };
        match (targets.first().copied(), connection.as_ref()) {
            (Some(site), Some(value)) => {
                validate_source_connection_target_attributes(site)?;
                if has_opaque_connection_target_content(&self.content, &nodes, site) {
                    return invalid(
                        "ODB connection change would discard opaque connection-target content",
                    );
                }
                let owner = connection_owner(&nodes, site);
                if owner.id != site.id
                    && matches!(value, Connection::Resource(_))
                    && has_opaque_connection_owner_content(&self.content, &nodes, owner, site)
                {
                    return invalid(
                        "ODB connection change would discard opaque database-description content",
                    );
                }
                if owner.id != site.id && !matches!(value, Connection::Resource(_)) {
                    let replacement = serialize_connection_target_for_site(
                        value,
                        &prefixes,
                        &self.content,
                        &nodes,
                        site,
                        site,
                    )?;
                    replace_span(&mut self.content, site.full.clone(), &replacement)?;
                } else {
                    let replacement = serialize_connection_owner_for_site(
                        value,
                        &prefixes,
                        &self.content,
                        &nodes,
                        site,
                        (owner.id != site.id).then_some(owner),
                    )?;
                    replace_span(&mut self.content, owner.full.clone(), &replacement)?;
                }
            },
            (Some(site), None) => {
                validate_source_connection_target_attributes(site)?;
                if has_opaque_connection_target_content(&self.content, &nodes, site) {
                    return invalid(
                        "ODB connection removal would discard opaque connection-target content",
                    );
                }
                let owner = connection_owner(&nodes, site);
                if owner.id != site.id
                    && has_opaque_connection_owner_content(&self.content, &nodes, owner, site)
                {
                    return invalid(
                        "ODB connection removal would discard opaque database-description content",
                    );
                }
                replace_span(&mut self.content, owner.full.clone(), "")?;
            },
            (None, Some(value)) => {
                if let Some(description) = unique_node(&nodes, "database-description")? {
                    if matches!(value, Connection::Resource(_)) {
                        if has_opaque_connection_owner_content_without_target(
                            &self.content,
                            &nodes,
                            description,
                        ) {
                            return invalid(
                                "ODB connection change would discard opaque database-description content",
                            );
                        }
                        let mut fragment = serialize_connection(value, &prefixes)?;
                        append_opaque_connection_attributes_for_scope(
                            &mut fragment,
                            &self.content,
                            &nodes,
                            description,
                            description,
                        )?;
                        replace_span(&mut self.content, description.full.clone(), &fragment)?;
                    } else {
                        let fragment = serialize_connection(value, &prefixes)?;
                        insert_child(&mut self.content, description, &fragment)?;
                    }
                } else if let Some(data) = unique_node(&nodes, "connection-data")? {
                    insert_child(
                        &mut self.content,
                        data,
                        &serialize_connection_owner(value, &prefixes)?,
                    )?;
                } else {
                    let source = unique_required_node(&nodes, "data-source")?;
                    let owner = serialize_connection_owner(value, &prefixes)?;
                    let wrapped = wrap_connection_data(&owner, &prefixes)?;
                    insert_child(&mut self.content, source, &wrapped)?;
                }
            },
            (None, None) => return invalid("ODB connection staging state is inconsistent"),
        }
        // Connection CRUD is published only when the resulting owner and all
        // modeled settings still satisfy their typed ODF projection. The
        // candidate is isolated by `stage_atomically`, so a malformed target
        // never leaks into the caller's edit.
        let _ = self.staged_catalog()?.settings()?;
        self.record(action, ChangeKind::Connection, "connection")
    }

    /// Replaces all modeled data-source settings in one source-bound,
    /// atomic operation.  Unknown attributes, producer children, and
    /// unrelated package members remain untouched.  The values are inert and
    /// no driver or credential is acquired.
    pub fn set_settings(&mut self, settings: DatabaseSettings) -> Result<()> {
        self.stage_atomically(|candidate| {
            validate_settings(&settings)?;
            candidate.set_settings_inner(&settings)
        })
    }

    /// Replaces only the inert login declaration, retaining the other
    /// settings families exactly as currently modeled.
    pub fn set_login_settings(&mut self, value: Option<LoginSettings>) -> Result<()> {
        let current = self.staged_catalog()?.settings()?.clone();
        self.set_settings(
            DatabaseSettings::new()
                .with_login(value)
                .with_driver(current.driver().cloned())
                .with_application_connection(current.application_connection().cloned()),
        )
    }

    /// Replaces only the inert driver-settings declaration.
    pub fn set_driver_settings(&mut self, value: Option<DriverSettings>) -> Result<()> {
        let current = self.staged_catalog()?.settings()?.clone();
        self.set_settings(
            DatabaseSettings::new()
                .with_login(current.login().cloned())
                .with_driver(value)
                .with_application_connection(current.application_connection().cloned()),
        )
    }

    /// Replaces only inert application/data-source filter settings.
    pub fn set_application_connection_settings(
        &mut self,
        value: Option<ApplicationConnectionSettings>,
    ) -> Result<()> {
        let current = self.staged_catalog()?.settings()?.clone();
        self.set_settings(
            DatabaseSettings::new()
                .with_login(current.login().cloned())
                .with_driver(current.driver().cloned())
                .with_application_connection(value),
        )
    }

    /// Alias for [`Edit::set_settings`].
    pub fn set_database_settings(&mut self, settings: DatabaseSettings) -> Result<()> {
        self.set_settings(settings)
    }

    fn set_settings_inner(&mut self, settings: &DatabaseSettings) -> Result<()> {
        let current = self.staged_catalog()?.settings()?.clone();
        if current == *settings {
            return Ok(());
        }
        let nodes = scan(&self.content)?;
        let source = unique_required_node(&nodes, "data-source")?;
        let connection_data =
            direct_child(&nodes, source, "connection-data")?.ok_or_else(|| {
                Error::InvalidFormat("ODB connection-data owner does not exist".to_string())
            })?;
        let prefixes = prefixes(&self.content)?;
        if current.login() != settings.login() {
            set_login_node(
                &mut self.content,
                &nodes,
                connection_data,
                settings.login(),
                &prefixes.database,
            )?;
        }
        if current.driver() != settings.driver() {
            let nodes = scan(&self.content)?;
            let source = unique_required_node(&nodes, "data-source")?;
            set_driver_node(
                &mut self.content,
                &nodes,
                source,
                settings.driver(),
                &prefixes.database,
            )?;
        }
        if current.application_connection() != settings.application_connection() {
            let nodes = scan(&self.content)?;
            let source = unique_required_node(&nodes, "data-source")?;
            set_application_node(
                &mut self.content,
                &nodes,
                source,
                settings.application_connection(),
                &prefixes.database,
            )?;
        }
        // Re-open the complete staged settings projection before publishing
        // the candidate. This catches required containers whose modeled
        // children were removed while opaque producer markup was retained.
        let published_catalog = self.staged_catalog()?;
        let published = published_catalog.settings()?;
        if published != settings {
            return invalid("ODB settings typed readback differs from authored values");
        }
        self.record(ChangeAction::Update, ChangeKind::Settings, "settings")
    }

    /// Adds an inert form or report component.
    pub fn add_component(&mut self, component: Component) -> Result<()> {
        let name = component.name().ok_or_else(|| {
            Error::InvalidFormat("ODB authored component requires a name".to_string())
        })?;
        validate_component(&component)?;
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
        self.record(
            ChangeAction::Create,
            ChangeKind::Component,
            &child_target(collection, name),
        )
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
        validate_component(&component)?;
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
            ChangeAction::Update,
            ChangeKind::Component,
            &child_target(component_collection(kind), name),
        )
    }

    /// Removes an inert form or report component.
    pub fn remove_component(&mut self, kind: ComponentKind, name: &str) -> Result<()> {
        let nodes = scan(&self.content)?;
        let site = find_component(&nodes, kind, name)?
            .ok_or_else(|| Error::InvalidFormat("ODB component does not exist".to_string()))?;
        replace_span(&mut self.content, site.full.clone(), "")?;
        self.record(
            ChangeAction::Remove,
            ChangeKind::Component,
            &child_target(component_collection(kind), name),
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
        let namespace = root.namespace_uri.as_deref().ok_or_else(|| {
            Error::InvalidFormat("ODB producer extension requires a namespace".to_string())
        })?;
        let nodes = scan(&self.content)?;
        let database = unique_required_node(&nodes, "database")?;
        if nodes.iter().any(|node| {
            node.parent == Some(database.id)
                && node.namespace_uri == root.namespace_uri
                && node.local == root.local
        }) {
            return invalid("ODB producer extension already exists");
        }
        insert_child(&mut self.content, database, xml)?;
        self.record(
            ChangeAction::Create,
            ChangeKind::ProducerExtension,
            &extension_target(namespace, &root.local),
        )
    }

    /// Replaces one unambiguous producer-extension subtree while retaining its
    /// expanded XML name.
    ///
    /// # Errors
    ///
    /// Returns an error for an absent/ambiguous selector, changed expanded
    /// name, noncompact XML, standard ODF root, malformed XML, or limit.
    pub fn replace_producer_extension(
        &mut self,
        namespace: &str,
        local: &str,
        xml: &str,
    ) -> Result<()> {
        validate_extension(xml)?;
        let replacement_nodes = scan(xml)?;
        let replacement = replacement_nodes
            .iter()
            .find(|node| node.parent.is_none())
            .ok_or_else(|| {
                Error::InvalidFormat("ODB producer extension has no root".to_string())
            })?;
        if replacement.namespace_uri.as_deref() != Some(namespace) || replacement.local != local {
            return invalid("ODB producer extension replacement changes its expanded name");
        }
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
        replace_span(&mut self.content, site.full.clone(), xml)?;
        self.record(
            ChangeAction::Update,
            ChangeKind::ProducerExtension,
            &extension_target(namespace, local),
        )
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
        self.record(
            ChangeAction::Remove,
            ChangeKind::ProducerExtension,
            &extension_target(namespace, local),
        )
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
        self.record(ChangeAction::Create, kind, &child_target(table, name))
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
        self.record(ChangeAction::Update, kind, &child_target(table, name))
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
        self.record(ChangeAction::Remove, kind, &child_target(table, name))
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

    fn record(&mut self, action: ChangeAction, kind: ChangeKind, target: &str) -> Result<()> {
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
            action,
            kind,
            target: target.to_owned(),
        });
        Ok(())
    }

    /// Atomically rebuilds, fully reopens, and semantically verifies the
    /// candidate package.
    pub fn commit(self) -> Result<Commit> {
        if self.content == self.source.content_xml()
            && self.payload_additions.is_empty()
            && self.payload_directories.is_empty()
        {
            return Ok(Commit::unchanged(self.source.clone()));
        }
        let protection = self.source.protection_status()?;
        if protection.is_signed()
            && matches!(
                self.policy.signature(),
                crate::SignaturePolicy::PreserveExactOnly
            )
        {
            return Err(Error::Unsupported(
                "ODB signature policy permits exact no-ops only".to_string(),
            ));
        }
        if protection.is_encrypted()
            && matches!(
                self.policy.encryption(),
                crate::EncryptionPolicy::PreserveExactOnly
            )
        {
            return Err(Error::Unsupported(
                "ODB encryption policy permits exact no-ops only".to_string(),
            ));
        }
        let typed_publication = self
            .changes
            .iter()
            .any(|change| matches!(change.kind, ChangeKind::Connection | ChangeKind::Settings));
        crate::codec::validate(&self.content)?;
        let snapshot = Database {
            package: self.source.package.rebuild_with_content_and_additions(
                &self.content,
                self.payload_additions,
                self.payload_directories,
            )?,
            settings: std::sync::Arc::new(std::sync::OnceLock::new()),
        };
        snapshot.catalog()?;
        if typed_publication {
            snapshot.settings()?;
        }
        if protection.is_signed()
            && matches!(
                self.policy.signature(),
                crate::SignaturePolicy::RemoveInvalidated
            )
            && snapshot.protection_status()?.is_signed()
        {
            return Err(Error::InvalidFormat(
                "ODB changed publication retained an invalidated signature member".to_string(),
            ));
        }
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

    /// Inventories the protection lifecycle of this exact publication.
    ///
    /// This reports signature/encryption metadata only; it does not verify a
    /// signature, request credentials, decrypt a member, or execute content.
    ///
    /// # Errors
    ///
    /// Returns an error if either package inventory cannot be inspected.
    pub fn protection_transition(&self) -> Result<crate::ProtectionTransition> {
        Ok(crate::ProtectionTransition::new(
            self.patch.source.protection_status()?,
            self.snapshot.protection_status()?,
        ))
    }

    #[must_use]
    pub fn into_database(self) -> Database {
        self.snapshot
    }
}

/// A byte-exact source-checked reversible ODB patch.
#[derive(Clone)]
pub struct Patch {
    pub(super) source: Database,
    pub(super) target: Database,
    pub(super) changes: Vec<Change>,
    pub(super) legacy_query: Option<QueryChange>,
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
    end_tag: Range<usize>,
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
    if source.len() > MAX_OUTPUT_BYTES {
        return invalid("ODB edit XML exceeds the byte limit");
    }
    let mut reader = NsReader::from_str(source);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut nodes = Vec::<Node>::new();
    let mut stack = Vec::<usize>::new();
    let mut events = 0usize;
    loop {
        events = events
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODB edit XML event count overflow".to_string()))?;
        if events > MAX_SCAN_NODES.saturating_mul(2).saturating_add(2) {
            return invalid("ODB edit XML event count exceeds the limit");
        }
        let (resolved, raw_event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODB edit XML: {error}")))?;
        let namespace = namespace_kind(&resolved);
        let namespace_uri = namespace_uri(&resolved)?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidFormat("ODB edit XML position exceeds this platform".to_string())
        })?;
        let start_event = matches!(&raw_event, Event::Start(_));
        match raw_event {
            Event::Start(element) | Event::Empty(element) => {
                if nodes.len() >= MAX_SCAN_NODES {
                    return invalid("ODB edit XML node count exceeds the limit");
                }
                let start = source[..end].rfind('<').ok_or_else(|| {
                    Error::InvalidFormat("ODB element start is missing".to_string())
                })?;
                let id = nodes.len();
                let local = std::str::from_utf8(element.local_name().as_ref())
                    .map_err(|_error| {
                        Error::InvalidFormat("ODB local name is not UTF-8".to_string())
                    })?
                    .to_owned();
                nodes.try_reserve(1).map_err(|source| Error::Allocation {
                    resource: "ODB edit XML nodes",
                    source,
                })?;
                nodes.push(Node {
                    id,
                    parent: stack.last().copied(),
                    namespace,
                    namespace_uri,
                    local,
                    start_tag: start..end,
                    end_tag: start..end,
                    full: start..end,
                    attrs: decode_attributes(&reader, &element)?,
                });
                if start_event {
                    if stack.len() >= MAX_XML_DEPTH {
                        return invalid("ODB edit XML nesting exceeds the limit");
                    }
                    stack.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "ODB edit XML stack",
                        source,
                    })?;
                    stack.push(id);
                }
            },
            Event::End(_) => {
                let start = source[..end].rfind('<').ok_or_else(|| {
                    Error::InvalidFormat("ODB element end is missing".to_string())
                })?;
                let id = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODB edit XML has an unmatched end tag".to_string())
                })?;
                nodes[id].end_tag = start..end;
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
    let mut count = 0usize;
    for raw in element.attributes() {
        count = count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODB edit attribute count overflow".to_string()))?;
        if count > MAX_SCAN_ATTRIBUTES {
            return invalid("ODB edit attribute count exceeds the limit");
        }
        let attribute =
            raw.map_err(|error| Error::InvalidFormat(format!("invalid ODB attribute: {error}")))?;
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let qname = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|_error| Error::InvalidFormat("ODB attribute name is not UTF-8".to_string()))?
            .to_owned();
        if qname.len() > MAX_VALUE_BYTES {
            return invalid("ODB edit attribute name exceeds the byte limit");
        }
        let local = std::str::from_utf8(local.as_ref())
            .map_err(|_error| {
                Error::InvalidFormat("ODB attribute local name is not UTF-8".to_string())
            })?
            .to_owned();
        let decoded = reader
            .decoder()
            .decode(attribute.value.as_ref())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODB attribute value: {error}"))
            })?;
        let value = quick_xml::escape::unescape(decoded.as_ref())
            .map(std::borrow::Cow::into_owned)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODB attribute value: {error}"))
            })?;
        if value.len() > MAX_VALUE_BYTES {
            return invalid("ODB edit attribute value exceeds the byte limit");
        }
        attributes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "ODB edit XML attributes",
                source,
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
        .ok_or_else(|| {
            Error::Unsupported("ODB editing requires a bound xlink namespace".to_string())
        })?
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
        let capacity = head
            .len()
            .checked_add(fragment.len())
            .and_then(|size| size.checked_add(name.len()))
            .and_then(|size| size.checked_add(5))
            .ok_or_else(|| Error::InvalidFormat("ODB XML edit size overflow".to_string()))?;
        if capacity > MAX_OUTPUT_BYTES {
            return invalid("ODB edited content exceeds the output limit");
        }
        let mut replacement = String::new();
        replacement
            .try_reserve(capacity)
            .map_err(|source| Error::Allocation {
                resource: "ODB XML edit replacement",
                source,
            })?;
        replacement.push_str(head);
        replacement.push('>');
        replacement.push_str(fragment);
        replacement.push_str("</");
        replacement.push_str(name);
        replacement.push('>');
        replace_span(source, owner.full.clone(), &replacement)
    } else {
        let closing_start = owner.end_tag.start;
        if owner.end_tag.end > source.len() || source.get(owner.end_tag.clone()).is_none() {
            return invalid("ODB owner closing tag span is invalid");
        }
        replace_span(source, closing_start..closing_start, fragment)
    }
}

fn insert_child_before_local(
    source: &mut String,
    owner: &Node,
    nodes: &[Node],
    later_locals: &[&str],
    fragment: &str,
) -> Result<()> {
    let before = nodes.iter().find(|node| {
        node.parent == Some(owner.id)
            && node.namespace == NamespaceKind::Database
            && later_locals.contains(&node.local.as_str())
    });
    if let Some(before) = before {
        replace_span(source, before.full.start..before.full.start, fragment)
    } else {
        insert_child(source, owner, fragment)
    }
}

fn direct_child<'a>(nodes: &'a [Node], owner: &'a Node, local: &str) -> Result<Option<&'a Node>> {
    unique_match(
        nodes
            .iter()
            .filter(|node| {
                node.parent == Some(owner.id)
                    && node.namespace == NamespaceKind::Database
                    && node.local == local
            })
            .collect(),
        "ODB settings child is ambiguous",
    )
}

fn set_login_node(
    source: &mut String,
    nodes: &[Node],
    owner: &Node,
    value: Option<&LoginSettings>,
    prefix: &str,
) -> Result<()> {
    let existing = direct_child(nodes, owner, "login")?;
    match (existing, value) {
        (Some(node), Some(value)) => {
            let edits = known_attribute_edits(
                source,
                node,
                prefix,
                [
                    ("user-name", value.user_name().map(str::to_owned)),
                    (
                        "use-system-user",
                        value
                            .use_system_user()
                            .map(|value| bool_text(value).to_owned()),
                    ),
                    (
                        "is-password-required",
                        value
                            .password_required()
                            .map(|value| bool_text(value).to_owned()),
                    ),
                    (
                        "login-timeout",
                        value.login_timeout().map(|value| value.to_string()),
                    ),
                ],
            )?;
            apply_edits(source, edits)
        },
        (Some(node), None) => {
            ensure_removable_subtree(
                source,
                nodes,
                node,
                &[
                    "user-name",
                    "use-system-user",
                    "is-password-required",
                    "login-timeout",
                ],
                &[],
            )?;
            replace_span(source, node.full.clone(), "")
        },
        (None, Some(value)) => {
            let fragment = serialize_login(value, prefix);
            insert_child(source, owner, &fragment)
        },
        (None, None) => Ok(()),
    }
}

fn set_driver_node(
    source: &mut String,
    nodes: &[Node],
    owner: &Node,
    value: Option<&DriverSettings>,
    prefix: &str,
) -> Result<()> {
    let existing = direct_child(nodes, owner, "driver-settings")?;
    match (existing, value) {
        (Some(node), None) => {
            ensure_removable_subtree(
                source,
                nodes,
                node,
                &[
                    "show-deleted",
                    "system-driver-settings",
                    "base-dn",
                    "is-first-row-header-line",
                    "parameter-name-substitution",
                    "additional-column-statement",
                    "row-retrieving-statement",
                    "field",
                    "string",
                    "decimal",
                    "thousand",
                    "encoding",
                ],
                &[
                    "auto-increment",
                    "delimiter",
                    "character-set",
                    "table-settings",
                    "table-setting",
                ],
            )?;
            replace_span(source, node.full.clone(), "")
        },
        (None, Some(value)) => insert_child_before_local(
            source,
            owner,
            nodes,
            &["application-connection-settings"],
            &serialize_driver(value, prefix),
        ),
        (None, None) => Ok(()),
        (Some(_), Some(value)) => {
            let nodes = scan(source)?;
            let owner = unique_required_node(&nodes, "data-source")?;
            let driver = direct_child(&nodes, owner, "driver-settings")?.ok_or_else(|| {
                Error::InvalidFormat("ODB driver-settings owner disappeared".to_string())
            })?;
            apply_edits(
                source,
                known_attribute_edits(
                    source,
                    driver,
                    prefix,
                    [
                        (
                            "show-deleted",
                            value
                                .show_deleted()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "system-driver-settings",
                            value.system_driver_settings().map(str::to_owned),
                        ),
                        ("base-dn", value.base_dn().map(str::to_owned)),
                        (
                            "is-first-row-header-line",
                            value
                                .first_row_header_line()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "parameter-name-substitution",
                            value
                                .parameter_name_substitution()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                    ],
                )?,
            )?;
            set_simple_driver_child(
                source,
                "auto-increment",
                value.auto_increment(),
                prefix,
                |value| {
                    vec![
                        (
                            "additional-column-statement",
                            value.additional_column_statement().map(str::to_owned),
                        ),
                        (
                            "row-retrieving-statement",
                            value.row_retrieving_statement().map(str::to_owned),
                        ),
                    ]
                },
                |value| serialize_auto_increment(value, prefix),
            )?;
            set_simple_driver_child(
                source,
                "delimiter",
                value.delimiter(),
                prefix,
                |value| {
                    vec![
                        ("field", value.field().map(str::to_owned)),
                        ("string", value.string().map(str::to_owned)),
                        ("decimal", value.decimal().map(str::to_owned)),
                        ("thousand", value.thousand().map(str::to_owned)),
                    ]
                },
                |value| serialize_delimiter(value, prefix),
            )?;
            set_simple_driver_child(
                source,
                "character-set",
                value.character_set(),
                prefix,
                |value| vec![("encoding", value.encoding().map(str::to_owned))],
                |value| serialize_character_set(value, prefix),
            )?;
            set_table_settings(source, value.table_settings(), prefix)
        },
    }
}

fn set_simple_driver_child<T>(
    source: &mut String,
    local: &str,
    value: Option<&T>,
    prefix: &str,
    attributes: impl FnOnce(&T) -> Vec<(&'static str, Option<String>)>,
    serialize: impl FnOnce(&T) -> String,
) -> Result<()> {
    let nodes = scan(source)?;
    let driver = unique_required_node(&nodes, "driver-settings")?;
    let existing = direct_child(&nodes, driver, local)?;
    match (existing, value) {
        (Some(node), Some(value)) => apply_edits(
            source,
            known_attribute_edits(source, node, prefix, attributes(value))?,
        ),
        (Some(node), None) => {
            ensure_removable_subtree(
                source,
                &nodes,
                node,
                &[
                    "additional-column-statement",
                    "row-retrieving-statement",
                    "field",
                    "string",
                    "decimal",
                    "thousand",
                    "encoding",
                ],
                &[],
            )?;
            replace_span(source, node.full.clone(), "")
        },
        (None, Some(value)) => {
            let later = match local {
                "auto-increment" => ["delimiter", "character-set", "table-settings"].as_slice(),
                "delimiter" => ["character-set", "table-settings"].as_slice(),
                "character-set" => ["table-settings"].as_slice(),
                _ => [].as_slice(),
            };
            insert_child_before_local(source, driver, &nodes, later, &serialize(value))
        },
        (None, None) => Ok(()),
    }
}

fn set_table_settings(source: &mut String, values: &[TableSetting], prefix: &str) -> Result<()> {
    let nodes = scan(source)?;
    let driver = unique_required_node(&nodes, "driver-settings")?;
    let existing = direct_child(&nodes, driver, "table-settings")?;
    match existing {
        Some(container) => {
            let children = direct_children(&nodes, container, "table-setting");
            let mut edits = Vec::new();
            for (child, value) in children.iter().zip(values.iter()) {
                edits.push(TextEdit {
                    range: child.full.clone(),
                    value: render_table_setting(source, &nodes, child, value, prefix)?,
                });
            }
            for child in children.iter().skip(values.len()) {
                ensure_removable_subtree(
                    source,
                    &nodes,
                    child,
                    &[
                        "is-first-row-header-line",
                        "show-deleted",
                        "field",
                        "string",
                        "decimal",
                        "thousand",
                        "encoding",
                    ],
                    &["delimiter", "character-set"],
                )?;
                edits.push(TextEdit {
                    range: child.full.clone(),
                    value: String::new(),
                });
            }
            if values.len() > children.len() {
                let mut fragment = String::new();
                for value in &values[children.len()..] {
                    fragment.push_str(&serialize_table_setting(value, prefix));
                }
                if container.full == container.start_tag {
                    let raw = source.get(container.full.clone()).ok_or_else(|| {
                        Error::InvalidFormat("ODB table-settings span is invalid".to_string())
                    })?;
                    let head = raw.strip_suffix("/>").ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODB empty table-settings syntax is invalid".to_string(),
                        )
                    })?;
                    let name = element_name(raw)?;
                    edits.push(TextEdit {
                        range: container.full.clone(),
                        value: format!("{head}>{fragment}</{name}>"),
                    });
                } else {
                    let insertion = container.end_tag.start;
                    if container.end_tag.end > source.len()
                        || source.get(container.end_tag.clone()).is_none()
                    {
                        return invalid("ODB table-settings closing tag span is invalid");
                    }
                    edits.push(TextEdit {
                        range: insertion..insertion,
                        value: fragment,
                    });
                }
            }
            if edits.is_empty() {
                return Ok(());
            }
            apply_edits(source, edits)?;
            if values.is_empty() {
                prune_empty_container(source, "table-settings", &["table-setting"])?;
            }
            Ok(())
        },
        None if values.is_empty() => Ok(()),
        None => {
            let fragment = values
                .iter()
                .map(|value| serialize_table_setting(value, prefix))
                .collect::<String>();
            insert_child(
                source,
                driver,
                &format!("<{prefix}:table-settings>{fragment}</{prefix}:table-settings>"),
            )
        },
    }
}

fn render_table_setting(
    source: &str,
    nodes: &[Node],
    node: &Node,
    value: &TableSetting,
    prefix: &str,
) -> Result<String> {
    let start = known_attribute_edits(
        source,
        node,
        prefix,
        [
            (
                "is-first-row-header-line",
                value
                    .first_row_header_line()
                    .map(|value| bool_text(value).to_owned()),
            ),
            (
                "show-deleted",
                value
                    .show_deleted()
                    .map(|value| bool_text(value).to_owned()),
            ),
        ],
    )?
    .pop()
    .ok_or_else(|| Error::InvalidFormat("ODB table-setting start span is missing".to_string()))?
    .value;
    let mut replacements = Vec::new();
    let delimiter = direct_child(nodes, node, "delimiter")?;
    if let Some(child) = delimiter {
        let replacement = render_simple_child(
            source,
            nodes,
            child,
            value.delimiter(),
            prefix,
            [
                (
                    "field",
                    value
                        .delimiter()
                        .and_then(|value| value.field())
                        .map(str::to_owned),
                ),
                (
                    "string",
                    value
                        .delimiter()
                        .and_then(|value| value.string())
                        .map(str::to_owned),
                ),
                (
                    "decimal",
                    value
                        .delimiter()
                        .and_then(|value| value.decimal())
                        .map(str::to_owned),
                ),
                (
                    "thousand",
                    value
                        .delimiter()
                        .and_then(|value| value.thousand())
                        .map(str::to_owned),
                ),
            ],
            &["field", "string", "decimal", "thousand"],
        )?;
        replacements.push((child.id, replacement));
    }
    let character_set = direct_child(nodes, node, "character-set")?;
    if let Some(child) = character_set {
        let replacement = render_simple_child(
            source,
            nodes,
            child,
            value.character_set(),
            prefix,
            [(
                "encoding",
                value
                    .character_set()
                    .and_then(|value| value.encoding())
                    .map(str::to_owned),
            )],
            &["encoding"],
        )?;
        replacements.push((child.id, replacement));
    }
    let mut insert = String::new();
    if delimiter.is_none() {
        if let Some(value) = value.delimiter() {
            insert.push_str(&serialize_delimiter(value, prefix));
        }
    }
    if character_set.is_none() {
        if let Some(value) = value.character_set() {
            insert.push_str(&serialize_character_set(value, prefix));
        }
    }
    let insert_before = if delimiter.is_none() && value.delimiter().is_some() {
        character_set.map(|child| child.id)
    } else {
        None
    };
    rebuild_element(
        source,
        nodes,
        node,
        &start,
        &replacements,
        &insert,
        insert_before,
    )
}

fn render_simple_child<T, I>(
    source: &str,
    nodes: &[Node],
    node: &Node,
    value: Option<&T>,
    prefix: &str,
    attributes: I,
    known_attrs: &[&str],
) -> Result<Option<String>>
where
    I: IntoIterator<Item = (&'static str, Option<String>)>,
{
    if value.is_none() {
        ensure_removable_subtree(source, nodes, node, known_attrs, &[])?;
        return Ok(None);
    }
    let start = known_attribute_edits(source, node, prefix, attributes)?
        .pop()
        .ok_or_else(|| {
            Error::InvalidFormat("ODB settings child start span is missing".to_string())
        })?
        .value;
    if node.full == node.start_tag {
        Ok(Some(start))
    } else {
        Ok(Some(format!(
            "{}{}",
            start,
            source
                .get(node.start_tag.end..node.full.end)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODB settings child body span is invalid".to_string())
                })?
        )))
    }
}

fn ensure_removable_subtree(
    source: &str,
    nodes: &[Node],
    node: &Node,
    known_attrs: &[&str],
    known_children: &[&str],
) -> Result<()> {
    let mut check = vec![node.id];
    while let Some(id) = check.pop() {
        let current = nodes.get(id).ok_or_else(|| {
            Error::InvalidFormat("ODB settings subtree span is invalid".to_string())
        })?;
        if current.namespace == NamespaceKind::Other
            || current.attrs.iter().any(|attribute| {
                attribute.namespace_uri.as_deref() != Some(str_from(DATABASE_NAMESPACE))
                    || !known_attrs.contains(&attribute.local.as_str())
            })
            || (current.full != current.start_tag
                && source
                    .get(current.start_tag.end..current.full.end)
                    .is_some_and(|inner| inner.contains("<!--") || inner.contains("<?")))
        {
            return Err(Error::Unsupported(
                "ODB settings edit would discard unknown XML".to_string(),
            ));
        }
        for child in nodes.iter().filter(|child| child.parent == Some(id)) {
            if child.namespace != NamespaceKind::Database
                || !known_children.contains(&child.local.as_str())
            {
                return Err(Error::Unsupported(
                    "ODB settings edit would discard unknown XML".to_string(),
                ));
            }
            check.push(child.id);
        }
    }
    Ok(())
}

fn rebuild_element(
    source: &str,
    nodes: &[Node],
    node: &Node,
    start: &str,
    replacements: &[(usize, Option<String>)],
    insert: &str,
    insert_before: Option<usize>,
) -> Result<String> {
    if node.full == node.start_tag {
        if insert.is_empty() {
            return Ok(start.to_owned());
        }
        let name = element_name(start)?;
        return Ok(format!(
            "{}>{insert}</{name}>",
            start.strip_suffix("/>").ok_or_else(|| {
                Error::InvalidFormat("ODB settings empty element syntax is invalid".to_string())
            })?
        ));
    }
    let closing_start = node.end_tag.start;
    if node.end_tag.end > source.len() || source.get(node.end_tag.clone()).is_none() {
        return invalid("ODB settings element closing tag span is invalid");
    }
    let mut output = start.to_owned();
    let mut cursor = node.start_tag.end;
    let mut inserted = false;
    for child in nodes.iter().filter(|child| child.parent == Some(node.id)) {
        output.push_str(source.get(cursor..child.full.start).ok_or_else(|| {
            Error::InvalidFormat("ODB settings element gap span is invalid".to_string())
        })?);
        if insert_before == Some(child.id) {
            output.push_str(insert);
            inserted = true;
        }
        if let Some((_, replacement)) = replacements.iter().find(|(id, _)| *id == child.id) {
            if let Some(replacement) = replacement {
                output.push_str(replacement);
            }
        } else {
            output.push_str(source.get(child.full.clone()).ok_or_else(|| {
                Error::InvalidFormat("ODB settings child span is invalid".to_string())
            })?);
        }
        cursor = child.full.end;
    }
    output.push_str(source.get(cursor..closing_start).ok_or_else(|| {
        Error::InvalidFormat("ODB settings element tail span is invalid".to_string())
    })?);
    if !inserted {
        output.push_str(insert);
    }
    output.push_str(source.get(closing_start..node.full.end).ok_or_else(|| {
        Error::InvalidFormat("ODB settings element closing span is invalid".to_string())
    })?);
    Ok(output)
}

fn element_name(start: &str) -> Result<&str> {
    start
        .get(1..)
        .and_then(|value| {
            value
                .split(|character: char| {
                    character.is_ascii_whitespace() || matches!(character, '/' | '>')
                })
                .next()
        })
        .ok_or_else(|| Error::InvalidFormat("ODB settings element name is missing".to_string()))
}

fn set_application_node(
    source: &mut String,
    nodes: &[Node],
    owner: &Node,
    value: Option<&ApplicationConnectionSettings>,
    prefix: &str,
) -> Result<()> {
    let existing = direct_child(nodes, owner, "application-connection-settings")?;
    match (existing, value) {
        (Some(node), None) => {
            ensure_removable_subtree(
                source,
                nodes,
                node,
                &[
                    "is-table-name-length-limited",
                    "enable-sql92-check",
                    "append-table-alias-name",
                    "ignore-driver-privileges",
                    "boolean-comparison-mode",
                    "use-catalog",
                    "max-row-count",
                    "suppress-version-columns",
                    "data-source-setting-name",
                    "data-source-setting-type",
                    "data-source-setting-is-list",
                ],
                &[
                    "table-filter",
                    "table-include-filter",
                    "table-exclude-filter",
                    "table-filter-pattern",
                    "table-type-filter",
                    "table-type",
                    "data-source-settings",
                    "data-source-setting",
                    "data-source-setting-value",
                ],
            )?;
            replace_span(source, node.full.clone(), "")
        },
        (None, Some(value)) => insert_child(source, owner, &serialize_application(value, prefix)),
        (None, None) => Ok(()),
        (Some(_), Some(value)) => {
            let nodes = scan(source)?;
            let owner = unique_required_node(&nodes, "data-source")?;
            let application = direct_child(&nodes, owner, "application-connection-settings")?
                .ok_or_else(|| {
                    Error::InvalidFormat("ODB application settings owner disappeared".to_string())
                })?;
            apply_edits(
                source,
                known_attribute_edits(
                    source,
                    application,
                    prefix,
                    [
                        (
                            "is-table-name-length-limited",
                            value
                                .table_name_length_limited()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "enable-sql92-check",
                            value
                                .enable_sql92_check()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "append-table-alias-name",
                            value
                                .append_table_alias_name()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "ignore-driver-privileges",
                            value
                                .ignore_driver_privileges()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "boolean-comparison-mode",
                            value
                                .boolean_comparison_mode()
                                .map(|value| value.as_str().to_owned()),
                        ),
                        (
                            "use-catalog",
                            value.use_catalog().map(|value| bool_text(value).to_owned()),
                        ),
                        (
                            "max-row-count",
                            value.max_row_count().map(|value| value.to_string()),
                        ),
                        (
                            "suppress-version-columns",
                            value
                                .suppress_version_columns()
                                .map(|value| bool_text(value).to_owned()),
                        ),
                    ],
                )?,
            )?;
            set_table_filter_node(source, value.table_filter(), prefix)?;
            set_table_type_filter_node(source, value.table_type_filter(), prefix)?;
            set_data_source_settings(source, value.data_source_settings(), prefix)
        },
    }
}

fn set_table_filter_node(
    source: &mut String,
    value: Option<&TableFilter>,
    prefix: &str,
) -> Result<()> {
    let nodes = scan(source)?;
    let application = unique_required_node(&nodes, "application-connection-settings")?;
    let existing = direct_child(&nodes, application, "table-filter")?;
    match (existing, value) {
        (None, Some(value)) => insert_child_before_local(
            source,
            application,
            &nodes,
            &["table-type-filter", "data-source-settings"],
            &serialize_table_filter(value, prefix),
        ),
        (None, None) => Ok(()),
        (Some(_), Some(value)) => {
            set_table_filter_group(source, "table-include-filter", value.include(), prefix)?;
            set_table_filter_group(source, "table-exclude-filter", value.exclude(), prefix)
        },
        (Some(_), None) => {
            set_table_filter_group(source, "table-include-filter", &[], prefix)?;
            set_table_filter_group(source, "table-exclude-filter", &[], prefix)?;
            prune_empty_container(
                source,
                "table-filter",
                &["table-include-filter", "table-exclude-filter"],
            )
        },
    }
}

fn set_table_filter_group(
    source: &mut String,
    local: &str,
    values: &[String],
    prefix: &str,
) -> Result<()> {
    let nodes = scan(source)?;
    let filter = unique_required_node(&nodes, "table-filter")?;
    let existing = direct_child(&nodes, filter, local)?;
    match existing {
        Some(group) => {
            set_repeated_text_children(source, group, "table-filter-pattern", values, prefix)?;
            if values.is_empty() {
                prune_empty_container(source, local, &["table-filter-pattern"])?;
            }
            Ok(())
        },
        None if values.is_empty() => Ok(()),
        None => {
            let element = if local == "table-include-filter" {
                "table-include-filter"
            } else {
                "table-exclude-filter"
            };
            let mut fragment = format!("<{prefix}:{element}>");
            for value in values {
                fragment.push_str(&serialize_text_element(
                    prefix,
                    "table-filter-pattern",
                    value,
                ));
            }
            fragment.push_str(&format!("</{prefix}:{element}>"));
            let later = if local == "table-include-filter" {
                ["table-exclude-filter"].as_slice()
            } else {
                [].as_slice()
            };
            insert_child_before_local(source, filter, &nodes, later, &fragment)
        },
    }
}

fn set_table_type_filter_node(
    source: &mut String,
    value: Option<&TableTypeFilter>,
    prefix: &str,
) -> Result<()> {
    let nodes = scan(source)?;
    let application = unique_required_node(&nodes, "application-connection-settings")?;
    let existing = direct_child(&nodes, application, "table-type-filter")?;
    match (existing, value) {
        (None, Some(value)) => insert_child_before_local(
            source,
            application,
            &nodes,
            &["data-source-settings"],
            &serialize_table_type_filter(value, prefix),
        ),
        (None, None) => Ok(()),
        (Some(filter), Some(value)) => {
            set_repeated_text_children(source, filter, "table-type", value.table_types(), prefix)
        },
        (Some(_), None) => {
            let nodes = scan(source)?;
            let filter = unique_required_node(&nodes, "table-type-filter")?;
            set_repeated_text_children(source, filter, "table-type", &[], prefix)?;
            prune_empty_container(source, "table-type-filter", &["table-type"])
        },
    }
}

fn set_repeated_text_children(
    source: &mut String,
    owner: &Node,
    local: &str,
    values: &[String],
    prefix: &str,
) -> Result<()> {
    let nodes = scan(source)?;
    let owner = nodes.get(owner.id).ok_or_else(|| {
        Error::InvalidFormat("ODB settings repeated-child owner span is invalid".to_string())
    })?;
    let children = direct_children(&nodes, owner, local);
    let mut edits = Vec::new();
    for (child, value) in children.iter().zip(values.iter()) {
        edits.push(TextEdit {
            range: child.full.clone(),
            value: text_element_replacement(source, &nodes, child, value)?,
        });
    }
    if children.len() > values.len() {
        for child in children[values.len()..].iter().rev() {
            ensure_removable_subtree(source, &nodes, child, &[], &[])?;
            edits.push(TextEdit {
                range: child.full.clone(),
                value: String::new(),
            });
        }
    }
    if values.len() > children.len() {
        let mut fragment = String::new();
        for value in &values[children.len()..] {
            fragment.push_str(&serialize_text_element(prefix, local, value));
        }
        if owner.full == owner.start_tag {
            let raw = source.get(owner.full.clone()).ok_or_else(|| {
                Error::InvalidFormat(
                    "ODB settings repeated-child owner span is invalid".to_string(),
                )
            })?;
            let head = raw.strip_suffix("/>").ok_or_else(|| {
                Error::InvalidFormat("ODB settings empty owner syntax is invalid".to_string())
            })?;
            let name = raw
                .get(1..)
                .and_then(|value| {
                    value
                        .split(|character: char| {
                            character.is_ascii_whitespace() || matches!(character, '/' | '>')
                        })
                        .next()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(
                        "ODB settings repeated-child owner name is missing".to_string(),
                    )
                })?;
            edits.push(TextEdit {
                range: owner.full.clone(),
                value: format!("{head}>{fragment}</{name}>"),
            });
        } else {
            let insertion = owner.end_tag.start;
            if owner.end_tag.end > source.len() || source.get(owner.end_tag.clone()).is_none() {
                return invalid("ODB settings repeated-child closing tag span is invalid");
            }
            edits.push(TextEdit {
                range: insertion..insertion,
                value: fragment,
            });
        }
    }
    apply_edits(source, edits)
}

fn direct_children<'a>(nodes: &'a [Node], owner: &Node, local: &str) -> Vec<&'a Node> {
    nodes
        .iter()
        .filter(|node| {
            node.parent == Some(owner.id)
                && node.namespace == NamespaceKind::Database
                && node.local == local
        })
        .collect()
}

fn serialize_text_element(prefix: &str, local: &str, value: &str) -> String {
    format!("<{prefix}:{local}>{}</{prefix}:{local}>", xml_escape(value))
}

fn text_element_replacement(
    source: &str,
    nodes: &[Node],
    node: &Node,
    value: &str,
) -> Result<String> {
    let start = source.get(node.start_tag.clone()).ok_or_else(|| {
        Error::InvalidFormat("ODB settings text start span is invalid".to_string())
    })?;
    if node.full == node.start_tag {
        let head = start.strip_suffix("/>").ok_or_else(|| {
            Error::InvalidFormat("ODB settings empty text element syntax is invalid".to_string())
        })?;
        let name = start
            .get(1..)
            .and_then(|value| {
                value
                    .split(|character: char| {
                        character.is_ascii_whitespace() || matches!(character, '/' | '>')
                    })
                    .next()
            })
            .ok_or_else(|| {
                Error::InvalidFormat("ODB settings text element name is missing".to_string())
            })?;
        return Ok(format!("{head}>{}</{name}>", xml_escape(value)));
    }
    let closing_start = node.end_tag.start;
    if node.end_tag.end > source.len() || source.get(node.end_tag.clone()).is_none() {
        return invalid("ODB settings text closing tag span is invalid");
    }
    let children = nodes
        .iter()
        .filter(|child| child.parent == Some(node.id))
        .collect::<Vec<_>>();
    let mut inner = String::new();
    let mut cursor = node.start_tag.end;
    let mut inserted = false;
    for child in children {
        let gap = source.get(cursor..child.full.start).ok_or_else(|| {
            Error::InvalidFormat("ODB settings text gap span is invalid".to_string())
        })?;
        inner.push_str(&replace_direct_text(gap, value, &mut inserted)?);
        inner.push_str(source.get(child.full.clone()).ok_or_else(|| {
            Error::InvalidFormat("ODB settings unknown child span is invalid".to_string())
        })?);
        cursor = child.full.end;
    }
    let gap = source.get(cursor..closing_start).ok_or_else(|| {
        Error::InvalidFormat("ODB settings text tail span is invalid".to_string())
    })?;
    inner.push_str(&replace_direct_text(gap, value, &mut inserted)?);
    if !inserted {
        inner.insert_str(0, &xml_escape(value));
    }
    Ok(format!(
        "{}{inner}{}",
        start,
        source.get(closing_start..node.full.end).ok_or_else(|| {
            Error::InvalidFormat("ODB settings text closing span is invalid".to_string())
        })?
    ))
}

fn replace_direct_text(gap: &str, value: &str, inserted: &mut bool) -> Result<String> {
    let mut reader = NsReader::from_str(gap);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut cursor = 0usize;
    let mut output = String::new();
    loop {
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODB setting text: {error}")))?;
        let end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            Error::InvalidFormat("ODB setting text position exceeds this platform".to_string())
        })?;
        match event {
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {
                if !*inserted {
                    output.push_str(&xml_escape(value));
                    *inserted = true;
                }
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return invalid("DOCTYPE is not permitted in ODB setting text");
            },
            Event::Start(_) | Event::End(_) => {
                output.push_str(gap.get(cursor..end).ok_or_else(|| {
                    Error::InvalidFormat("ODB setting text event span is invalid".to_string())
                })?);
            },
            Event::Empty(_) | Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {
                output.push_str(gap.get(cursor..end).ok_or_else(|| {
                    Error::InvalidFormat("ODB setting lexical span is invalid".to_string())
                })?);
            },
        }
        cursor = end;
        buffer.clear();
    }
    output.push_str(gap.get(cursor..).ok_or_else(|| {
        Error::InvalidFormat("ODB setting text tail span is invalid".to_string())
    })?);
    Ok(output)
}

fn prune_empty_container(source: &mut String, local: &str, _known_children: &[&str]) -> Result<()> {
    let nodes = scan(source)?;
    let owner = unique_required_node(&nodes, local)?;
    let has_unknown_attribute = !owner.attrs.is_empty();
    let has_child = nodes.iter().any(|node| node.parent == Some(owner.id));
    let has_unknown_lexical = if owner.full == owner.start_tag {
        false
    } else {
        source
            .get(owner.start_tag.end..owner.full.end)
            .is_some_and(|inner| inner.contains("<!--") || inner.contains("<?"))
    };
    if !has_unknown_attribute && !has_child && !has_unknown_lexical {
        replace_span(source, owner.full.clone(), "")?;
    }
    Ok(())
}

fn set_data_source_settings(
    source: &mut String,
    values: &[DataSourceSetting],
    prefix: &str,
) -> Result<()> {
    let nodes = scan(source)?;
    let application = unique_required_node(&nodes, "application-connection-settings")?;
    let existing = direct_child(&nodes, application, "data-source-settings")?;
    match existing {
        Some(container) => {
            let children = direct_children(&nodes, container, "data-source-setting");
            let mut edits = Vec::new();
            for (child, value) in children.iter().zip(values.iter()) {
                edits.push(TextEdit {
                    range: child.full.clone(),
                    value: render_data_source_setting(source, &nodes, child, value, prefix)?,
                });
            }
            for child in children.iter().skip(values.len()) {
                ensure_removable_subtree(
                    source,
                    &nodes,
                    child,
                    &[
                        "data-source-setting-name",
                        "data-source-setting-type",
                        "data-source-setting-is-list",
                    ],
                    &["data-source-setting-value"],
                )?;
                edits.push(TextEdit {
                    range: child.full.clone(),
                    value: String::new(),
                });
            }
            if values.len() > children.len() {
                let mut fragment = String::new();
                for value in &values[children.len()..] {
                    fragment.push_str(&serialize_data_source_setting(value, prefix));
                }
                if container.full == container.start_tag {
                    let raw = source.get(container.full.clone()).ok_or_else(|| {
                        Error::InvalidFormat("ODB data-source-settings span is invalid".to_string())
                    })?;
                    let head = raw.strip_suffix("/>").ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODB empty data-source-settings syntax is invalid".to_string(),
                        )
                    })?;
                    let name = element_name(raw)?;
                    edits.push(TextEdit {
                        range: container.full.clone(),
                        value: format!("{head}>{fragment}</{name}>"),
                    });
                } else {
                    let insertion = container.end_tag.start;
                    if container.end_tag.end > source.len()
                        || source.get(container.end_tag.clone()).is_none()
                    {
                        return invalid("ODB data-source-settings closing tag span is invalid");
                    }
                    edits.push(TextEdit {
                        range: insertion..insertion,
                        value: fragment,
                    });
                }
            }
            if edits.is_empty() {
                return Ok(());
            }
            apply_edits(source, edits)?;
            if values.is_empty() {
                prune_empty_container(source, "data-source-settings", &["data-source-setting"])?;
            }
            Ok(())
        },
        None if values.is_empty() => Ok(()),
        None => insert_child(
            source,
            application,
            &format!(
                "<{prefix}:data-source-settings>{}</{prefix}:data-source-settings>",
                values
                    .iter()
                    .map(|value| serialize_data_source_setting(value, prefix))
                    .collect::<String>()
            ),
        ),
    }
}

fn render_data_source_setting(
    source: &str,
    nodes: &[Node],
    node: &Node,
    value: &DataSourceSetting,
    prefix: &str,
) -> Result<String> {
    let start = known_attribute_edits(
        source,
        node,
        prefix,
        [
            ("data-source-setting-name", Some(value.name().to_owned())),
            (
                "data-source-setting-type",
                Some(value.setting_type().as_str().to_owned()),
            ),
            (
                "data-source-setting-is-list",
                value.is_list().map(|value| bool_text(value).to_owned()),
            ),
        ],
    )?
    .pop()
    .ok_or_else(|| {
        Error::InvalidFormat("ODB data-source-setting start span is missing".to_string())
    })?
    .value;
    let children = direct_children(nodes, node, "data-source-setting-value");
    let mut replacements = Vec::new();
    for (child, text) in children.iter().zip(value.values().iter()) {
        replacements.push((
            child.id,
            Some(text_element_replacement(source, nodes, child, text)?),
        ));
    }
    for child in children.iter().skip(value.values().len()) {
        ensure_removable_subtree(source, nodes, child, &[], &[])?;
        replacements.push((child.id, None));
    }
    let mut insert = String::new();
    for text in &value.values()[children.len().min(value.values().len())..] {
        insert.push_str(&serialize_text_element(
            prefix,
            "data-source-setting-value",
            text,
        ));
    }
    rebuild_element(source, nodes, node, &start, &replacements, &insert, None)
}

fn known_attribute_edits<'a, I>(
    source: &str,
    node: &Node,
    prefix: &str,
    attributes: I,
) -> Result<Vec<TextEdit>>
where
    I: IntoIterator<Item = (&'a str, Option<String>)>,
{
    let mut raw = source
        .get(node.start_tag.clone())
        .ok_or_else(|| Error::InvalidFormat("ODB settings start-tag span is invalid".to_string()))?
        .to_owned();
    for (local, value) in attributes {
        let existing = node.attrs.iter().find(|attribute| {
            attribute.namespace_uri.as_deref() == Some(str_from(DATABASE_NAMESPACE))
                && attribute.local == local
        });
        raw = match (existing, value) {
            (Some(attribute), Some(value)) => {
                replace_attribute(&raw, &attribute.qname, &xml_attribute_escape(&value))?
            },
            (Some(attribute), None) => remove_attribute(&raw, &attribute.qname)?,
            (None, Some(value)) => insert_attribute(
                &raw,
                &format!("{prefix}:{local}"),
                &xml_attribute_escape(&value),
            )?,
            (None, None) => raw,
        };
    }
    Ok(vec![TextEdit {
        range: node.start_tag.clone(),
        value: raw,
    }])
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
            replace_attribute(raw, &attribute.qname, &xml_attribute_escape(value))?
        },
        (Some(attribute), None) => remove_attribute(raw, &attribute.qname)?,
        (None, Some(value)) => insert_attribute(
            raw,
            &format!("{prefix}:{local}"),
            &xml_attribute_escape(value),
        )?,
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
    let removed = range.end - range.start;
    let additional = value.len().saturating_sub(removed);
    source
        .try_reserve(additional)
        .map_err(|source| Error::Allocation {
            resource: "ODB edited content",
            source,
        })?;
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
    if let Some(command) = table.filter_statement() {
        children.push_str(&empty_element(
            prefix,
            "filter-statement",
            vec![("command", command.to_owned())],
        ));
    }
    if let Some(command) = table.order_statement() {
        children.push_str(&empty_element(
            prefix,
            "order-statement",
            vec![("command", command.to_owned())],
        ));
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
    if let Some(value) = column.default_value() {
        attrs.push(("default-value", value.to_owned()));
    }
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
    let mut children = String::new();
    if !query.columns().is_empty() {
        children.push_str(&format!("<{prefix}:columns>"));
        for column in query.columns() {
            children.push_str(&serialize_column(column, false, prefix));
        }
        children.push_str(&format!("</{prefix}:columns>"));
    }
    if let Some(command) = query.filter_statement() {
        children.push_str(&empty_element(
            prefix,
            "filter-statement",
            vec![("command", command.to_owned())],
        ));
    }
    if let Some(command) = query.order_statement() {
        children.push_str(&empty_element(
            prefix,
            "order-statement",
            vec![("command", command.to_owned())],
        ));
    }
    if let Some(target) = query.update_target() {
        let mut update = vec![("name", target.name().to_owned())];
        if let Some(value) = target.schema_name() {
            update.push(("schema-name", value.to_owned()));
        }
        if let Some(value) = target.catalog_name() {
            update.push(("catalog-name", value.to_owned()));
        }
        children.push_str(&empty_element(prefix, "update-table", update));
    }
    element_with_children(prefix, "query", attrs, &children)
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
            xml_attribute_escape(value),
            prefixes.xlink
        ));
    }
    output.push_str("/>");
    output
}

fn serialize_login(value: &LoginSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    if let Some(value) = value.user_name() {
        attrs.push(("user-name", value.to_owned()));
    }
    push_bool(&mut attrs, "use-system-user", value.use_system_user());
    push_bool(
        &mut attrs,
        "is-password-required",
        value.password_required(),
    );
    if let Some(value) = value.login_timeout() {
        attrs.push(("login-timeout", value.to_string()));
    }
    empty_element(prefix, "login", attrs)
}

fn serialize_auto_increment(value: &AutoIncrementSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    if let Some(value) = value.additional_column_statement() {
        attrs.push(("additional-column-statement", value.to_owned()));
    }
    if let Some(value) = value.row_retrieving_statement() {
        attrs.push(("row-retrieving-statement", value.to_owned()));
    }
    empty_element(prefix, "auto-increment", attrs)
}

fn serialize_delimiter(value: &DelimiterSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    for (name, value) in [
        ("field", value.field()),
        ("string", value.string()),
        ("decimal", value.decimal()),
        ("thousand", value.thousand()),
    ] {
        if let Some(value) = value {
            attrs.push((name, value.to_owned()));
        }
    }
    empty_element(prefix, "delimiter", attrs)
}

fn serialize_character_set(value: &CharacterSetSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    if let Some(value) = value.encoding() {
        attrs.push(("encoding", value.to_owned()));
    }
    empty_element(prefix, "character-set", attrs)
}

fn serialize_table_setting(value: &TableSetting, prefix: &str) -> String {
    let mut attrs = Vec::new();
    push_bool(
        &mut attrs,
        "is-first-row-header-line",
        value.first_row_header_line(),
    );
    push_bool(&mut attrs, "show-deleted", value.show_deleted());
    let mut children = String::new();
    if let Some(value) = value.delimiter() {
        children.push_str(&serialize_delimiter(value, prefix));
    }
    if let Some(value) = value.character_set() {
        children.push_str(&serialize_character_set(value, prefix));
    }
    element_with_children(prefix, "table-setting", attrs, &children)
}

fn serialize_driver(value: &DriverSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    push_bool(&mut attrs, "show-deleted", value.show_deleted());
    if let Some(value) = value.system_driver_settings() {
        attrs.push(("system-driver-settings", value.to_owned()));
    }
    if let Some(value) = value.base_dn() {
        attrs.push(("base-dn", value.to_owned()));
    }
    push_bool(
        &mut attrs,
        "is-first-row-header-line",
        value.first_row_header_line(),
    );
    push_bool(
        &mut attrs,
        "parameter-name-substitution",
        value.parameter_name_substitution(),
    );
    let mut children = String::new();
    if let Some(value) = value.auto_increment() {
        children.push_str(&serialize_auto_increment(value, prefix));
    }
    if let Some(value) = value.delimiter() {
        children.push_str(&serialize_delimiter(value, prefix));
    }
    if let Some(value) = value.character_set() {
        children.push_str(&serialize_character_set(value, prefix));
    }
    if !value.table_settings().is_empty() {
        children.push_str(&format!("<{prefix}:table-settings>"));
        for setting in value.table_settings() {
            children.push_str(&serialize_table_setting(setting, prefix));
        }
        children.push_str(&format!("</{prefix}:table-settings>"));
    }
    element_with_children(prefix, "driver-settings", attrs, &children)
}

fn serialize_table_filter(value: &TableFilter, prefix: &str) -> String {
    let mut children = String::new();
    if !value.include().is_empty() {
        children.push_str(&format!("<{prefix}:table-include-filter>"));
        for pattern in value.include() {
            children.push_str(&format!(
                "<{prefix}:table-filter-pattern>{}</{prefix}:table-filter-pattern>",
                xml_escape(pattern)
            ));
        }
        children.push_str(&format!("</{prefix}:table-include-filter>"));
    }
    if !value.exclude().is_empty() {
        children.push_str(&format!("<{prefix}:table-exclude-filter>"));
        for pattern in value.exclude() {
            children.push_str(&format!(
                "<{prefix}:table-filter-pattern>{}</{prefix}:table-filter-pattern>",
                xml_escape(pattern)
            ));
        }
        children.push_str(&format!("</{prefix}:table-exclude-filter>"));
    }
    element_with_children(
        prefix,
        "table-filter",
        Vec::<(&'static str, String)>::new(),
        &children,
    )
}

fn serialize_table_type_filter(value: &TableTypeFilter, prefix: &str) -> String {
    let mut children = String::new();
    for table_type in value.table_types() {
        children.push_str(&format!(
            "<{prefix}:table-type>{}</{prefix}:table-type>",
            xml_escape(table_type)
        ));
    }
    element_with_children(
        prefix,
        "table-type-filter",
        Vec::<(&'static str, String)>::new(),
        &children,
    )
}

fn serialize_data_source_setting(value: &DataSourceSetting, prefix: &str) -> String {
    let mut attrs = vec![
        ("data-source-setting-name", value.name().to_owned()),
        (
            "data-source-setting-type",
            value.setting_type().as_str().to_owned(),
        ),
    ];
    push_bool(&mut attrs, "data-source-setting-is-list", value.is_list());
    let mut children = String::new();
    for setting_value in value.values() {
        children.push_str(&format!(
            "<{prefix}:data-source-setting-value>{}</{prefix}:data-source-setting-value>",
            xml_escape(setting_value)
        ));
    }
    element_with_children(prefix, "data-source-setting", attrs, &children)
}

fn serialize_application(value: &ApplicationConnectionSettings, prefix: &str) -> String {
    let mut attrs = Vec::new();
    push_bool(
        &mut attrs,
        "is-table-name-length-limited",
        value.table_name_length_limited(),
    );
    push_bool(&mut attrs, "enable-sql92-check", value.enable_sql92_check());
    push_bool(
        &mut attrs,
        "append-table-alias-name",
        value.append_table_alias_name(),
    );
    push_bool(
        &mut attrs,
        "ignore-driver-privileges",
        value.ignore_driver_privileges(),
    );
    if let Some(value) = value.boolean_comparison_mode() {
        attrs.push(("boolean-comparison-mode", value.as_str().to_owned()));
    }
    push_bool(&mut attrs, "use-catalog", value.use_catalog());
    if let Some(value) = value.max_row_count() {
        attrs.push(("max-row-count", value.to_string()));
    }
    push_bool(
        &mut attrs,
        "suppress-version-columns",
        value.suppress_version_columns(),
    );
    let mut children = String::new();
    if let Some(value) = value.table_filter() {
        children.push_str(&serialize_table_filter(value, prefix));
    }
    if let Some(value) = value.table_type_filter() {
        children.push_str(&serialize_table_type_filter(value, prefix));
    }
    if !value.data_source_settings().is_empty() {
        children.push_str(&format!("<{prefix}:data-source-settings>"));
        for setting in value.data_source_settings() {
            children.push_str(&serialize_data_source_setting(setting, prefix));
        }
        children.push_str(&format!("</{prefix}:data-source-settings>"));
    }
    element_with_children(prefix, "application-connection-settings", attrs, &children)
}

fn serialize_connection(connection: &Connection, prefixes: &Prefixes) -> Result<String> {
    let mut output = String::new();
    output
        .try_reserve(connection_output_capacity(connection, prefixes)?)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection XML",
            source,
        })?;
    match connection {
        Connection::File(href) => {
            write_xml(
                &mut output,
                format_args!(
                    "<{0}:file-based-database {1}:type=\"simple\" {1}:href=\"{2}\" {0}:media-type=\"application/octet-stream\"/>",
                    prefixes.database,
                    prefixes.xlink,
                    xml_attribute_escape(href)
                ),
            )?;
        },
        Connection::FileTarget(target) => write_file_target(&mut output, target, prefixes)?,
        Connection::Resource(href) => {
            write_xml(
                &mut output,
                format_args!(
                    "<{0}:connection-resource {1}:href=\"{2}\" {1}:type=\"simple\"/>",
                    prefixes.database,
                    prefixes.xlink,
                    xml_attribute_escape(href)
                ),
            )?;
        },
        Connection::Server { host, database } => {
            write_xml(
                &mut output,
                format_args!(
                    "<{0}:server-database {0}:type=\"sdbc:generic\" {0}:hostname=\"{1}\" {0}:database-name=\"{2}\"/>",
                    prefixes.database,
                    xml_attribute_escape(host),
                    xml_attribute_escape(database)
                ),
            )?;
        },
        Connection::ServerTarget(target) => write_server_target(&mut output, target, prefixes)?,
    }
    Ok(output)
}

fn serialize_connection_owner(connection: &Connection, prefixes: &Prefixes) -> Result<String> {
    let target = serialize_connection(connection, prefixes)?;
    if matches!(connection, Connection::Resource(_)) {
        Ok(target)
    } else {
        let mut output = String::new();
        let overhead = prefixes
            .database
            .len()
            .checked_mul(2)
            .and_then(|value| value.checked_add(128))
            .and_then(|value| value.checked_add(target.len()))
            .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))?;
        if overhead > MAX_OUTPUT_BYTES {
            return invalid("ODB connection XML exceeds the output limit");
        }
        output
            .try_reserve(overhead)
            .map_err(|source| Error::Allocation {
                resource: "ODB connection owner XML",
                source,
            })?;
        write_xml(
            &mut output,
            format_args!("<{0}:database-description>", prefixes.database),
        )?;
        output.push_str(&target);
        write_xml(
            &mut output,
            format_args!("</{0}:database-description>", prefixes.database),
        )?;
        Ok(output)
    }
}

fn serialize_connection_owner_for_site(
    connection: &Connection,
    prefixes: &Prefixes,
    source: &str,
    nodes: &[Node],
    target: &Node,
    description: Option<&Node>,
) -> Result<String> {
    let scope = description.unwrap_or(target);
    let mut target_xml =
        serialize_connection_target_for_site(connection, prefixes, source, nodes, target, scope)?;
    if matches!(connection, Connection::Resource(_)) {
        if let Some(description) = description {
            append_opaque_connection_attributes_for_scope(
                &mut target_xml,
                source,
                nodes,
                description,
                scope,
            )?;
        }
        return Ok(target_xml);
    }

    let overhead = prefixes
        .database
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(128))
        .and_then(|value| value.checked_add(target_xml.len()))
        .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))?;
    if overhead > MAX_OUTPUT_BYTES {
        return invalid("ODB connection XML exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve(overhead)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection owner XML",
            source,
        })?;
    write_xml(
        &mut output,
        format_args!("<{0}:database-description", prefixes.database),
    )?;
    if let Some(description) = description {
        append_opaque_connection_attributes_for_scope(
            &mut output,
            source,
            nodes,
            description,
            scope,
        )?;
    }
    output.push('>');
    output.push_str(&target_xml);
    write_xml(
        &mut output,
        format_args!("</{0}:database-description>", prefixes.database),
    )?;
    Ok(output)
}

fn serialize_connection_target_for_site(
    connection: &Connection,
    prefixes: &Prefixes,
    source: &str,
    nodes: &[Node],
    target: &Node,
    scope: &Node,
) -> Result<String> {
    let mut target_xml = serialize_connection(connection, prefixes)?;
    append_required_connection_namespace_bindings(&mut target_xml, nodes, target, scope, prefixes)?;
    append_opaque_connection_attributes_for_scope(&mut target_xml, source, nodes, target, scope)?;
    Ok(target_xml)
}

fn append_required_connection_namespace_bindings(
    output: &mut String,
    nodes: &[Node],
    target: &Node,
    scope: &Node,
    prefixes: &Prefixes,
) -> Result<()> {
    let mut declarations = Vec::new();
    let mut current = Some(target.id);
    while let Some(id) = current {
        let current_node = nodes.get(id).ok_or_else(|| {
            Error::InvalidFormat("ODB connection namespace owner is out of bounds".to_string())
        })?;
        if id == scope.id {
            break;
        }
        current = current_node.parent;
    }
    if current != Some(scope.id) {
        return invalid("ODB connection namespace scope is not an ancestor");
    }

    for (prefix, namespace) in [
        (prefixes.database.as_str(), str_from(DATABASE_NAMESPACE)),
        (prefixes.xlink.as_str(), str_from(XLINK_NAMESPACE)),
    ] {
        let declaration_name = format!("xmlns:{prefix}");
        if output.contains(&format!(" {declaration_name}=")) {
            continue;
        }
        let mut found = None;
        let mut current = Some(target.id);
        while let Some(id) = current {
            let current_node = nodes.get(id).ok_or_else(|| {
                Error::InvalidFormat("ODB connection namespace owner is out of bounds".to_string())
            })?;
            if let Some(attribute) = current_node
                .attrs
                .iter()
                .find(|attribute| attribute.qname == declaration_name)
            {
                if attribute.value != namespace {
                    return invalid("ODB connection namespace binding changed in source");
                }
                found = Some(attribute);
                break;
            }
            if id == scope.id {
                break;
            }
            current = current_node.parent;
        }
        if let Some(attribute) = found {
            declarations.push(attribute);
        }
    }
    if declarations.is_empty() {
        return Ok(());
    }
    let capacity = declarations.iter().try_fold(0usize, |total, attribute| {
        let escaped = xml_attribute_escaped_len(&attribute.value)?;
        total
            .checked_add(attribute.qname.len())
            .and_then(|value| value.checked_add(escaped))
            .and_then(|value| value.checked_add(4))
            .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))
    })?;
    if output
        .len()
        .checked_add(capacity)
        .is_none_or(|size| size > MAX_OUTPUT_BYTES)
    {
        return invalid("ODB connection XML exceeds the output limit");
    }
    output
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection namespace declarations",
            source,
        })?;
    let mut attributes = String::new();
    attributes
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection namespace declarations",
            source,
        })?;
    for attribute in declarations {
        write_xml(
            &mut attributes,
            format_args!(
                " {0}=\"{1}\"",
                attribute.qname,
                xml_attribute_escape(&attribute.value)
            ),
        )?;
    }
    if let Some(position) = output.rfind("/>") {
        output.insert_str(position, &attributes);
    } else {
        output.push_str(&attributes);
    }
    Ok(())
}

fn append_opaque_connection_attributes_for_scope(
    output: &mut String,
    source: &str,
    nodes: &[Node],
    node: &Node,
    scope: &Node,
) -> Result<()> {
    let attributes = collect_opaque_connection_attributes(source, nodes, node, scope, output)?;
    if attributes.is_empty() {
        return Ok(());
    }
    if output
        .len()
        .checked_add(attributes.len())
        .is_none_or(|size| size > MAX_OUTPUT_BYTES)
    {
        return invalid("ODB connection XML exceeds the output limit");
    }
    output
        .try_reserve(attributes.len())
        .map_err(|source| Error::Allocation {
            resource: "ODB connection extension attributes",
            source,
        })?;
    if let Some(position) = output.rfind("/>") {
        output.insert_str(position, &attributes);
    } else {
        output.push_str(&attributes);
    }
    Ok(())
}

fn collect_opaque_connection_attributes(
    _source: &str,
    nodes: &[Node],
    node: &Node,
    scope: &Node,
    existing: &str,
) -> Result<String> {
    let opaque = node
        .attrs
        .iter()
        .filter(|attribute| is_opaque_connection_attribute(attribute))
        .collect::<Vec<_>>();
    if opaque.is_empty() {
        return Ok(String::new());
    }

    let mut prefixes = BTreeSet::new();
    for attribute in &opaque {
        if let Some((prefix, _)) = attribute.qname.split_once(':')
            && prefix != "xml"
        {
            prefixes.insert(prefix.to_owned());
        }
    }

    let mut declarations = Vec::<&Attr>::new();
    let mut declared = BTreeSet::new();
    let mut current = Some(node.id);
    while let Some(id) = current {
        let current_node = nodes.get(id).ok_or_else(|| {
            Error::InvalidFormat("ODB connection namespace owner is out of bounds".to_string())
        })?;
        for attribute in &current_node.attrs {
            let Some(prefix) = attribute.qname.strip_prefix("xmlns:") else {
                continue;
            };
            if prefixes.contains(prefix)
                && declared.insert(prefix.to_owned())
                && !existing.contains(&format!(" {}=\"", attribute.qname))
            {
                declarations.push(attribute);
            }
        }
        if id == scope.id {
            break;
        }
        current = current_node.parent;
    }
    if current != Some(scope.id) {
        return invalid("ODB connection namespace scope is not an ancestor");
    }

    for attribute in &opaque {
        if let Some((prefix, _)) = attribute.qname.split_once(':')
            && prefix != "xml"
            && attribute.namespace_uri.is_none()
            && !declared.contains(prefix)
        {
            return invalid("ODB extension attribute prefix is not bound in source");
        }
    }

    let mut output = String::new();
    let capacity =
        declarations
            .iter()
            .chain(opaque.iter())
            .try_fold(0usize, |total, attribute| {
                let escaped = xml_attribute_escaped_len(&attribute.value)?;
                total
                    .checked_add(attribute.qname.len())
                    .and_then(|value| value.checked_add(escaped))
                    .and_then(|value| value.checked_add(4))
                    .ok_or_else(|| {
                        Error::InvalidFormat("ODB connection XML size overflow".to_string())
                    })
            })?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODB connection XML exceeds the output limit");
    }
    output
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection extension attributes",
            source,
        })?;
    for attribute in declarations.into_iter().chain(opaque) {
        write_xml(
            &mut output,
            format_args!(
                " {0}=\"{1}\"",
                attribute.qname,
                xml_attribute_escape(&attribute.value)
            ),
        )?;
    }
    Ok(output)
}

fn is_opaque_connection_attribute(attribute: &Attr) -> bool {
    !attribute.qname.eq("xmlns")
        && !attribute.qname.starts_with("xmlns:")
        && !matches!(
            attribute.namespace_uri.as_deref(),
            Some(namespace)
                if namespace == str_from(DATABASE_NAMESPACE)
                    || namespace == str_from(XLINK_NAMESPACE)
                    || namespace == str_from(OFFICE_NAMESPACE)
        )
}

fn wrap_connection_data(owner: &str, prefixes: &Prefixes) -> Result<String> {
    let capacity = prefixes
        .database
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(owner.len()))
        .and_then(|value| value.checked_add(128))
        .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))?;
    if capacity > MAX_OUTPUT_BYTES {
        return invalid("ODB connection XML exceeds the output limit");
    }
    let mut output = String::new();
    output
        .try_reserve(capacity)
        .map_err(|source| Error::Allocation {
            resource: "ODB connection-data XML",
            source,
        })?;
    write_xml(
        &mut output,
        format_args!("<{0}:connection-data>", prefixes.database),
    )?;
    output.push_str(owner);
    write_xml(
        &mut output,
        format_args!("</{0}:connection-data>", prefixes.database),
    )?;
    Ok(output)
}

fn write_file_target(
    output: &mut String,
    target: &FileDatabaseTarget,
    prefixes: &Prefixes,
) -> Result<()> {
    write_xml(
        output,
        format_args!(
            "<{0}:file-based-database {1}:type=\"simple\" {1}:href=\"{2}\" {0}:media-type=\"{3}\"",
            prefixes.database,
            prefixes.xlink,
            xml_attribute_escape(target.href()),
            xml_attribute_escape(target.media_type())
        ),
    )?;
    if let Some(extension) = target.extension() {
        write_xml(
            output,
            format_args!(
                " {0}:extension=\"{1}\"",
                prefixes.database,
                xml_attribute_escape(extension)
            ),
        )?;
    }
    output.push_str("/>");
    Ok(())
}

fn write_server_target(
    output: &mut String,
    target: &ServerDatabaseTarget,
    prefixes: &Prefixes,
) -> Result<()> {
    write_xml(
        output,
        format_args!(
            "<{0}:server-database {0}:type=\"{1}\"",
            prefixes.database,
            xml_attribute_escape(target.database_type())
        ),
    )?;
    if let Some(namespace) = target.database_type_namespace() {
        let prefix = target
            .database_type()
            .split_once(':')
            .map(|(prefix, _)| prefix)
            .ok_or_else(|| {
                Error::InvalidFormat("ODB server connection type is invalid".to_string())
            })?;
        write_xml(
            output,
            format_args!(" xmlns:{prefix}=\"{}\"", xml_attribute_escape(namespace)),
        )?;
    }
    match target.address() {
        Some(ServerDatabaseAddress::Host { hostname, port }) => {
            write_xml(
                output,
                format_args!(
                    " {0}:hostname=\"{1}\"",
                    prefixes.database,
                    xml_attribute_escape(hostname)
                ),
            )?;
            if let Some(port) = port {
                write_xml(
                    output,
                    format_args!(" {0}:port=\"{port}\"", prefixes.database),
                )?;
            }
        },
        Some(ServerDatabaseAddress::LocalSocket(local_socket)) => {
            write_xml(
                output,
                format_args!(
                    " {0}:local-socket=\"{1}\"",
                    prefixes.database,
                    xml_attribute_escape(local_socket)
                ),
            )?;
        },
        None => {},
    }
    if let Some(database_name) = target.database_name() {
        write_xml(
            output,
            format_args!(
                " {0}:database-name=\"{1}\"",
                prefixes.database,
                xml_attribute_escape(database_name)
            ),
        )?;
    }
    output.push_str("/>");
    Ok(())
}

fn write_xml(output: &mut String, value: std::fmt::Arguments<'_>) -> Result<()> {
    std::fmt::Write::write_fmt(output, value)
        .map_err(|_| Error::InvalidFormat("ODB connection XML formatting failed".to_string()))
}

fn connection_output_capacity(connection: &Connection, prefixes: &Prefixes) -> Result<usize> {
    let fields = match connection {
        Connection::File(href) | Connection::Resource(href) => [href.len(), 0, 0, 0],
        Connection::FileTarget(target) => [
            target.href().len(),
            target.media_type().len(),
            target.extension().map_or(0, str::len),
            0,
        ],
        Connection::Server { host, database } => [host.len(), database.len(), 0, 0],
        Connection::ServerTarget(target) => {
            let address = target.address().map_or(0, |address| match address {
                ServerDatabaseAddress::Host { hostname, .. } => hostname.len(),
                ServerDatabaseAddress::LocalSocket(value) => value.len(),
            });
            [
                target.database_type().len(),
                address,
                target.database_name().map_or(0, str::len),
                0,
            ]
        },
    };
    let prefix_bytes = prefixes
        .database
        .len()
        .checked_add(prefixes.xlink.len())
        .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))?
        .checked_mul(6)
        .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))?;
    fields
        .into_iter()
        .map(|value| {
            value
                .checked_mul(6)
                .ok_or_else(|| Error::InvalidFormat("ODB connection XML size overflow".to_string()))
        })
        .try_fold(
            512usize.checked_add(prefix_bytes).ok_or_else(|| {
                Error::InvalidFormat("ODB connection XML size overflow".to_string())
            })?,
            |total, value| {
                total.checked_add(value?).ok_or_else(|| {
                    Error::InvalidFormat("ODB connection XML size overflow".to_string())
                })
            },
        )
        .and_then(|capacity| {
            if capacity > MAX_OUTPUT_BYTES {
                Err(Error::InvalidFormat(
                    "ODB connection XML exceeds the output limit".to_string(),
                ))
            } else {
                Ok(capacity)
            }
        })
}

fn connection_owner<'a>(nodes: &'a [Node], target: &'a Node) -> &'a Node {
    target
        .parent
        .map(|id| &nodes[id])
        .filter(|parent| parent.local == "database-description")
        .unwrap_or(target)
}

fn validate_source_connection_target_attributes(target: &Node) -> Result<()> {
    for attribute in &target.attrs {
        if attribute.qname == "xmlns" || attribute.qname.starts_with("xmlns:") {
            continue;
        }
        let allowed = match attribute.namespace_uri.as_deref() {
            Some(namespace) if namespace == str_from(DATABASE_NAMESPACE) => {
                match target.local.as_str() {
                    "connection-resource" => false,
                    "file-based-database" => {
                        matches!(attribute.local.as_str(), "media-type" | "extension")
                    },
                    "server-database" => matches!(
                        attribute.local.as_str(),
                        "type" | "hostname" | "port" | "local-socket" | "database-name"
                    ),
                    _ => false,
                }
            },
            Some(namespace) if namespace == str_from(XLINK_NAMESPACE) => {
                match target.local.as_str() {
                    "connection-resource" => matches!(
                        attribute.local.as_str(),
                        "type" | "href" | "show" | "actuate"
                    ),
                    "file-based-database" => matches!(attribute.local.as_str(), "type" | "href"),
                    "server-database" => false,
                    _ => false,
                }
            },
            Some(namespace) if namespace == str_from(OFFICE_NAMESPACE) => false,
            None if reserved_prefix_string(&attribute.qname) => false,
            Some(_) | None => true,
        };
        if !allowed {
            return invalid("ODB connection target contains an unknown reserved attribute");
        }
    }
    Ok(())
}

fn reserved_prefix_string(qname: &str) -> bool {
    qname
        .split_once(':')
        .is_some_and(|(prefix, _)| matches!(prefix, "db" | "office" | "xlink"))
}

fn has_opaque_connection_owner_content(
    source: &str,
    nodes: &[Node],
    owner: &Node,
    target: &Node,
) -> bool {
    if nodes
        .iter()
        .any(|node| node.parent == Some(owner.id) && node.id != target.id)
    {
        return true;
    }
    let Some(target_before) = source.get(owner.start_tag.end..target.full.start) else {
        return true;
    };
    let Some(target_after) = source.get(target.full.end..owner.end_tag.start) else {
        return true;
    };
    !target_before.bytes().all(|byte| byte.is_ascii_whitespace())
        || !target_after.bytes().all(|byte| byte.is_ascii_whitespace())
}

fn has_opaque_connection_target_content(source: &str, nodes: &[Node], target: &Node) -> bool {
    if nodes.iter().any(|node| node.parent == Some(target.id)) {
        return true;
    }
    if target.full == target.start_tag {
        return false;
    }
    source
        .get(target.start_tag.end..target.end_tag.start)
        .is_none_or(|body| !body.is_empty())
}

fn has_opaque_connection_owner_content_without_target(
    source: &str,
    nodes: &[Node],
    owner: &Node,
) -> bool {
    if nodes.iter().any(|node| node.parent == Some(owner.id)) {
        return true;
    }
    if owner.full == owner.start_tag {
        return false;
    }
    let Some(body) = source.get(owner.start_tag.end..owner.end_tag.start) else {
        return true;
    };
    !body.bytes().all(|byte| byte.is_ascii_whitespace())
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
        output.push_str(&xml_attribute_escape(&value));
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

struct Dependent<'a> {
    node: &'a Node,
    table: &'a str,
    name: &'a str,
    kind: ChangeKind,
}

fn column_dependents<'a>(
    nodes: &'a [Node],
    table: &'a Node,
    table_name: &'a str,
    name: &str,
) -> Vec<Dependent<'a>> {
    let mut dependents = Vec::new();
    for owner in descendants(nodes, table).filter(|node| {
        matches!(node.local.as_str(), "key" | "index")
            && descendants(nodes, node).any(|column| {
                matches!(column.local.as_str(), "key-column" | "index-column")
                    && attribute(column, DATABASE_NAMESPACE, "name")
                        .is_some_and(|value| value == name)
            })
    }) {
        if let Some(owner_name) = attribute(owner, DATABASE_NAMESPACE, "name") {
            dependents.push(Dependent {
                node: owner,
                table: table_name,
                name: owner_name,
                kind: if owner.local == "key" {
                    ChangeKind::Key
                } else {
                    ChangeKind::Index
                },
            });
        }
    }
    for key in nodes.iter().filter(|node| {
        node.local == "key"
            && attribute(node, DATABASE_NAMESPACE, "referenced-table-name")
                .is_some_and(|value| value == table_name)
            && descendants(nodes, node).any(|column| {
                column.local == "key-column"
                    && attribute(column, DATABASE_NAMESPACE, "related-column-name")
                        .is_some_and(|value| value == name)
            })
    }) {
        let owner_table = ancestors(nodes, key)
            .find(|node| node.local == "table-definition")
            .and_then(|node| attribute(node, DATABASE_NAMESPACE, "name"));
        if let (Some(owner_table), Some(key_name)) =
            (owner_table, attribute(key, DATABASE_NAMESPACE, "name"))
            && !dependents
                .iter()
                .any(|dependent| dependent.node.id == key.id)
        {
            dependents.push(Dependent {
                node: key,
                table: owner_table,
                name: key_name,
                kind: ChangeKind::Key,
            });
        }
    }
    dependents
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

fn owned_table<'catalog>(
    catalog: &'catalog crate::OwnedCatalog,
    name: &str,
) -> Result<&'catalog Table> {
    let mut matches = catalog.tables().iter().filter(|table| table.name() == name);
    let table = matches
        .next()
        .ok_or_else(|| Error::InvalidFormat("ODB transfer table does not exist".to_string()))?;
    if matches.next().is_some() {
        return invalid("ODB transfer table selector is ambiguous");
    }
    Ok(table)
}

fn owned_query<'catalog>(
    catalog: &'catalog crate::OwnedCatalog,
    name: &str,
) -> Result<&'catalog Query> {
    let mut matches = catalog
        .queries()
        .iter()
        .filter(|query| query.name() == name);
    let query = matches
        .next()
        .ok_or_else(|| Error::InvalidFormat("ODB transfer query does not exist".to_string()))?;
    if matches.next().is_some() {
        return invalid("ODB transfer query selector is ambiguous");
    }
    Ok(query)
}

fn validate_source_table_dependencies(catalog: &crate::OwnedCatalog, table: &Table) -> Result<()> {
    for key in table.keys() {
        validate_key_columns(table, key)?;
        if let Some(referenced) = key.referenced_table() {
            let target = owned_table(catalog, referenced)?;
            validate_related_columns(target, key)?;
        }
    }
    validate_table_indices(table)
}

fn validate_key_columns(table: &Table, key: &Key) -> Result<()> {
    if key.columns().iter().any(|key_column| {
        key_column
            .name()
            .is_none_or(|name| !has_column(table, name))
    }) {
        return invalid("ODB key references a missing local column");
    }
    Ok(())
}

fn validate_related_columns(table: &Table, key: &Key) -> Result<()> {
    if key.columns().iter().any(|key_column| {
        key_column
            .related_column()
            .is_some_and(|name| !has_column(table, name))
    }) {
        return invalid("ODB key references a missing related column");
    }
    Ok(())
}

fn validate_table_indices(table: &Table) -> Result<()> {
    if table.indices().iter().any(|index| {
        index
            .columns()
            .iter()
            .any(|index_column| !has_column(table, index_column.name()))
    }) {
        return invalid("ODB index references a missing local column");
    }
    Ok(())
}

fn has_column(table: &Table, name: &str) -> bool {
    table.columns().iter().any(|column| column.name() == name)
}

fn dependency_refusal<T>(message: &str) -> Result<T> {
    Err(Error::Unsupported(message.to_owned()))
}

fn extension_target(namespace: &str, local: &str) -> String {
    child_target(namespace, local)
}

fn child_target(owner: &str, name: &str) -> String {
    format!("{}:{owner}{name}", owner.len())
}

fn validate_table(table: &Table) -> Result<()> {
    validate_name(table.name(), "table")?;
    if table.kind() == TableKind::Representation
        && (!table.keys().is_empty() || !table.indices().is_empty())
    {
        return invalid("ODB table representations cannot author schema keys or indices");
    }
    if table.kind() == TableKind::Definition
        && (table.filter_statement().is_some() || table.order_statement().is_some())
    {
        return invalid("ODB table definitions cannot author presentation statements");
    }
    for (kind, value) in [
        ("table filter statement", table.filter_statement()),
        ("table order statement", table.order_statement()),
    ] {
        if let Some(value) = value {
            validate_value(value, kind)?;
        }
    }
    let mut column_names = BTreeSet::new();
    for column in table.columns() {
        validate_column(column)?;
        if !column_names.insert(column.name()) {
            return invalid("ODB table contains duplicate column names");
        }
    }
    let mut key_names = BTreeSet::new();
    let mut primary_keys = 0usize;
    for key in table.keys() {
        validate_key(key)?;
        let name = key
            .name()
            .ok_or_else(|| Error::InvalidFormat("ODB authored key requires a name".to_string()))?;
        if !key_names.insert(name) {
            return invalid("ODB table contains duplicate key names");
        }
        if key.kind() == crate::KeyKind::Primary {
            primary_keys += 1;
            if primary_keys > 1 {
                return invalid("ODB table contains multiple primary keys");
            }
        }
        for mapping in key.columns() {
            if let Some(name) = mapping.name()
                && !table.columns().iter().any(|column| column.name() == name)
            {
                return invalid("ODB key maps an absent local column");
            }
        }
    }
    let mut index_names = BTreeSet::new();
    for index in table.indices() {
        validate_index(index)?;
        if !index_names.insert(index.name()) {
            return invalid("ODB table contains duplicate index names");
        }
        if index.columns().iter().any(|indexed| {
            !table
                .columns()
                .iter()
                .any(|column| column.name() == indexed.name())
        }) {
            return invalid("ODB index maps an absent local column");
        }
    }
    Ok(())
}

fn validate_column(column: &Column) -> Result<()> {
    validate_name(column.name(), "column")?;
    if let Some(value) = column.type_name() {
        validate_value(value, "column type name")?;
    }
    if let Some(value) = column.default_value() {
        validate_value(value, "column default value")?;
    }
    if column.precision() == Some(0) || column.scale() == Some(0) {
        return invalid("ODB column precision and scale must be positive");
    }
    Ok(())
}

fn validate_key(key: &Key) -> Result<()> {
    let name = key
        .name()
        .ok_or_else(|| Error::InvalidFormat("ODB authored key requires a name".to_string()))?;
    validate_name(name, "key")?;
    if key.columns().is_empty() {
        return invalid("ODB authored key requires at least one column");
    }
    if key.kind() == crate::KeyKind::Foreign && key.referenced_table().is_none() {
        return invalid("ODB foreign key requires a referenced table");
    }
    if key.kind() != crate::KeyKind::Foreign
        && (key.referenced_table().is_some()
            || key.update_rule().is_some()
            || key.delete_rule().is_some())
    {
        return invalid("ODB non-foreign key cannot carry relation metadata");
    }
    if let Some(value) = key.referenced_table() {
        validate_name(value, "referenced table")?;
    }
    let mut columns = BTreeSet::new();
    for column in key.columns() {
        let local = column.name().ok_or_else(|| {
            Error::InvalidFormat("ODB authored key column requires a name".to_string())
        })?;
        validate_name(local, "key column")?;
        if !columns.insert(local) {
            return invalid("ODB key contains duplicate local column mappings");
        }
        if let Some(value) = column.related_column() {
            validate_name(value, "related key column")?;
            if key.kind() != crate::KeyKind::Foreign {
                return invalid("ODB non-foreign key cannot map related columns");
            }
        }
    }
    Ok(())
}

fn validate_index(index: &Index) -> Result<()> {
    validate_name(index.name(), "index")?;
    if index.columns().is_empty() {
        return invalid("ODB authored index requires at least one column");
    }
    let mut columns = BTreeSet::new();
    for column in index.columns() {
        validate_name(column.name(), "index column")?;
        if !columns.insert(column.name()) {
            return invalid("ODB index contains duplicate column mappings");
        }
    }
    Ok(())
}

fn validate_connection(connection: &Connection) -> Result<()> {
    match connection {
        Connection::File(href) | Connection::Resource(href) => {
            validate_value(href, "connection target")
        },
        Connection::FileTarget(target) => {
            validate_value(target.href(), "file connection href")?;
            validate_value(target.media_type(), "file connection media type")?;
            if let Some(extension) = target.extension() {
                validate_value(extension, "file connection extension")?;
            }
            Ok(())
        },
        Connection::Server { host, database } => {
            validate_value(host, "connection host")?;
            validate_value(database, "connection database")?;
            Err(Error::InvalidFormat(
                "ODB legacy server connection cannot be authored without a bound db:type QName"
                    .to_string(),
            ))
        },
        Connection::ServerTarget(target) => {
            validate_namespaced_token(target.database_type(), "server connection type")?;
            if target.database_type().split_once(':').is_none() {
                return invalid("ODB server connection type is invalid");
            }
            if let Some(namespace) = target.database_type_namespace() {
                validate_value(namespace, "server connection type namespace")?;
                if namespace.is_empty() {
                    return invalid("ODB server connection type namespace is empty");
                }
            } else {
                return invalid("ODB server connection type prefix is not bound");
            }
            match target.address() {
                Some(ServerDatabaseAddress::Host { hostname, port }) => {
                    validate_value(hostname, "connection host")?;
                    if port == &Some(0) {
                        return invalid("ODB server connection port must be positive");
                    }
                },
                Some(ServerDatabaseAddress::LocalSocket(local_socket)) => {
                    validate_value(local_socket, "connection local socket")?;
                },
                None => {},
            }
            if let Some(database) = target.database_name() {
                validate_value(database, "connection database")?;
            }
            Ok(())
        },
    }
}

fn validate_namespaced_token(value: &str, kind: &str) -> Result<()> {
    let mut parts = value.split(':');
    let prefix = parts.next().unwrap_or_default();
    let local = parts.next().unwrap_or_default();
    if !is_ncname(prefix) || !is_ncname(local) || parts.next().is_some() {
        return invalid(&format!("invalid ODB {kind}"));
    }
    validate_value(value, kind)
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

#[derive(Default)]
struct SettingsBudget {
    value_count: usize,
    escaped_bytes: usize,
}

impl SettingsBudget {
    fn item(&mut self, kind: &str) -> Result<()> {
        self.add(1, 128, kind)
    }

    fn lexical(&mut self, bytes: usize, kind: &str) -> Result<()> {
        self.add(1, bytes.saturating_add(64), kind)
    }

    fn text(&mut self, value: &str, kind: &str) -> Result<()> {
        validate_value(value, kind)?;
        let escaped = xml_escaped_len(value)?;
        self.add(1, escaped.saturating_add(64), kind)
    }

    fn attribute(&mut self, value: &str, kind: &str) -> Result<()> {
        validate_value(value, kind)?;
        let escaped = xml_attribute_escaped_len(value)?;
        self.add(1, escaped.saturating_add(64), kind)
    }

    fn optional_bool(&mut self, value: Option<bool>, kind: &str) -> Result<()> {
        if value.is_some() {
            self.lexical(5, kind)?;
        }
        Ok(())
    }

    fn optional_i64(&mut self, value: Option<i64>, kind: &str) -> Result<()> {
        if value.is_some() {
            self.lexical(20, kind)?;
        }
        Ok(())
    }

    fn add(&mut self, values: usize, bytes: usize, kind: &str) -> Result<()> {
        self.value_count = self
            .value_count
            .checked_add(values)
            .ok_or_else(|| Error::InvalidFormat("ODB settings value count overflow".to_string()))?;
        if self.value_count > MAX_SETTINGS_VALUE_COUNT {
            return Err(Error::InvalidFormat(format!(
                "ODB settings total value count exceeds {MAX_SETTINGS_VALUE_COUNT} ({kind})"
            )));
        }
        self.escaped_bytes = self
            .escaped_bytes
            .checked_add(bytes)
            .ok_or_else(|| Error::InvalidFormat("ODB settings output size overflow".to_string()))?;
        if self.escaped_bytes > MAX_SETTINGS_OUTPUT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "ODB settings escaped output exceeds {MAX_SETTINGS_OUTPUT_BYTES} bytes ({kind})"
            )));
        }
        Ok(())
    }
}

fn validate_settings(settings: &DatabaseSettings) -> Result<()> {
    let mut budget = SettingsBudget::default();
    if let Some(login) = settings.login() {
        budget.item("login")?;
        if login.user_name().is_some() && login.use_system_user().is_some() {
            return invalid("ODB login cannot contain both user-name and use-system-user");
        }
        if let Some(value) = login.user_name() {
            budget.attribute(value, "login user-name")?;
        }
        budget.optional_bool(login.use_system_user(), "use-system-user")?;
        budget.optional_bool(login.password_required(), "is-password-required")?;
        if login.login_timeout() == Some(0) {
            return invalid("ODB login-timeout must be positive");
        }
        if login.login_timeout().is_some() {
            budget.lexical(20, "login-timeout")?;
        }
    }
    if let Some(driver) = settings.driver() {
        budget.item("driver-settings")?;
        budget.optional_bool(driver.show_deleted(), "show-deleted")?;
        for (kind, value) in [
            ("system-driver-settings", driver.system_driver_settings()),
            ("base-dn", driver.base_dn()),
        ] {
            if let Some(value) = value {
                budget.attribute(value, kind)?;
            }
        }
        budget.optional_bool(driver.first_row_header_line(), "is-first-row-header-line")?;
        budget.optional_bool(
            driver.parameter_name_substitution(),
            "parameter-name-substitution",
        )?;
        if let Some(value) = driver.auto_increment() {
            validate_auto_increment(value, &mut budget)?;
        }
        if let Some(value) = driver.delimiter() {
            validate_delimiter(value, &mut budget)?;
        }
        if let Some(value) = driver.character_set() {
            validate_character_set(value, &mut budget)?;
        }
        if driver.table_settings().len() > MAX_TABLE_SETTINGS {
            return invalid("ODB driver table-settings exceed the table-setting limit");
        }
        if !driver.table_settings().is_empty() {
            budget.item("table-settings")?;
        }
        for value in driver.table_settings() {
            validate_table_setting(value, &mut budget)?;
        }
    }
    if let Some(application) = settings.application_connection() {
        budget.item("application-connection-settings")?;
        budget.optional_bool(
            application.table_name_length_limited(),
            "is-table-name-length-limited",
        )?;
        budget.optional_bool(application.enable_sql92_check(), "enable-sql92-check")?;
        budget.optional_bool(
            application.append_table_alias_name(),
            "append-table-alias-name",
        )?;
        budget.optional_bool(
            application.ignore_driver_privileges(),
            "ignore-driver-privileges",
        )?;
        if let Some(value) = application.boolean_comparison_mode() {
            budget.lexical(value.as_str().len(), "boolean-comparison-mode")?;
        }
        budget.optional_bool(application.use_catalog(), "use-catalog")?;
        budget.optional_i64(application.max_row_count(), "max-row-count")?;
        budget.optional_bool(
            application.suppress_version_columns(),
            "suppress-version-columns",
        )?;
        if let Some(value) = application.table_filter() {
            budget.item("table-filter")?;
            if !value.include().is_empty() {
                budget.item("table-include-filter")?;
            }
            for pattern in value.include() {
                budget.text(pattern, "table filter pattern")?;
            }
            if !value.exclude().is_empty() {
                budget.item("table-exclude-filter")?;
            }
            for pattern in value.exclude() {
                budget.text(pattern, "table filter pattern")?;
            }
        }
        if let Some(value) = application.table_type_filter() {
            budget.item("table-type-filter")?;
            for table_type in value.table_types() {
                budget.text(table_type, "table type filter")?;
            }
        }
        if application.data_source_settings().len() > MAX_OPERATIONS {
            return invalid("ODB data-source settings exceed the operation limit");
        }
        if !application.data_source_settings().is_empty() {
            budget.item("data-source-settings")?;
        }
        for setting in application.data_source_settings() {
            budget.item("data-source-setting")?;
            validate_value(setting.name(), "data-source setting name")?;
            budget.attribute(setting.name(), "data-source setting name")?;
            budget.lexical(
                setting.setting_type().as_str().len(),
                "data-source setting type",
            )?;
            budget.optional_bool(setting.is_list(), "data-source-setting-is-list")?;
            if setting.values().is_empty() {
                return invalid("ODB data-source setting requires at least one value");
            }
            for value in setting.values() {
                budget.text(value, "data-source setting value")?;
            }
        }
    }
    Ok(())
}

fn validate_auto_increment(
    value: &AutoIncrementSettings,
    budget: &mut SettingsBudget,
) -> Result<()> {
    budget.item("auto-increment")?;
    for (kind, text) in [
        (
            "additional-column-statement",
            value.additional_column_statement(),
        ),
        ("row-retrieving-statement", value.row_retrieving_statement()),
    ] {
        if let Some(text) = text {
            budget.attribute(text, kind)?;
        }
    }
    Ok(())
}

fn validate_delimiter(value: &DelimiterSettings, budget: &mut SettingsBudget) -> Result<()> {
    budget.item("delimiter")?;
    for (kind, text) in [
        ("delimiter field", value.field()),
        ("delimiter string", value.string()),
        ("delimiter decimal", value.decimal()),
        ("delimiter thousand", value.thousand()),
    ] {
        if let Some(text) = text {
            budget.attribute(text, kind)?;
        }
    }
    Ok(())
}

fn validate_character_set(value: &CharacterSetSettings, budget: &mut SettingsBudget) -> Result<()> {
    budget.item("character-set")?;
    value.validate()?;
    if let Some(value) = value.encoding() {
        budget.attribute(value, "character-set encoding")?;
    }
    Ok(())
}

fn validate_table_setting(value: &TableSetting, budget: &mut SettingsBudget) -> Result<()> {
    budget.item("table-setting")?;
    budget.optional_bool(value.first_row_header_line(), "is-first-row-header-line")?;
    budget.optional_bool(value.show_deleted(), "show-deleted")?;
    if let Some(value) = value.delimiter() {
        validate_delimiter(value, budget)?;
    }
    if let Some(value) = value.character_set() {
        validate_character_set(value, budget)?;
    }
    Ok(())
}

fn validate_component(component: &Component) -> Result<()> {
    let name = component.name().ok_or_else(|| {
        Error::InvalidFormat("ODB authored component requires a name".to_string())
    })?;
    validate_name(name, "component")?;
    for (kind, value) in [
        ("component title", component.title()),
        ("component description", component.description()),
        ("component href", component.href()),
    ] {
        if let Some(value) = value {
            validate_value(value, kind)?;
        }
    }
    Ok(())
}

fn validate_query(query: &Query) -> Result<()> {
    validate_name(query.name(), "query")?;
    validate_value(query.command(), "query command")?;
    for column in query.columns() {
        validate_column(column)?;
    }
    for (kind, value) in [
        ("query filter statement", query.filter_statement()),
        ("query order statement", query.order_statement()),
    ] {
        if let Some(value) = value {
            validate_value(value, kind)?;
        }
    }
    if let Some(target) = query.update_target() {
        validate_name(target.name(), "query update table")?;
        for (kind, value) in [
            ("query update schema", target.schema_name()),
            ("query update catalog", target.catalog_name()),
        ] {
            if let Some(value) = value {
                validate_value(value, kind)?;
            }
        }
    }
    Ok(())
}

fn local_component_prefix(href: &str) -> Result<Option<String>> {
    if href.contains(':') || href.starts_with('#') {
        return Ok(None);
    }
    if href.is_empty()
        || href.starts_with('/')
        || href.contains('\\')
        || href.contains('?')
        || href.contains('#')
    {
        return invalid("ODB component package path is not a safe relative path");
    }
    let path = href.trim_end_matches('/');
    if path.is_empty()
        || path
            .split('/')
            .any(|segment| segment.is_empty() || matches!(segment, "." | ".."))
    {
        return invalid("ODB component package path is not a safe relative path");
    }
    Ok(Some(format!("{path}/")))
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
    let escaped = quick_xml::escape::escape(value);
    let mut output = String::new();
    output.reserve(escaped.len());
    for character in escaped.chars() {
        if character == '\r' {
            output.push_str("&#xD;");
        } else {
            output.push(character);
        }
    }
    output
}

fn xml_attribute_escape(value: &str) -> String {
    let mut output = String::new();
    for character in quick_xml::escape::escape(value).chars() {
        match character {
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            _ => output.push(character),
        }
    }
    output
}

fn xml_escaped_len(value: &str) -> Result<usize> {
    value.bytes().try_fold(0usize, |length, byte| {
        let escaped = match byte {
            b'\r' => 5,
            b'&' => 5,
            b'<' | b'>' => 4,
            b'"' | b'\'' => 6,
            _ => 1,
        };
        length
            .checked_add(escaped)
            .ok_or_else(|| Error::InvalidFormat("ODB settings output size overflow".to_string()))
    })
}

fn xml_attribute_escaped_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let escaped = match character {
            '\t' | '\n' | '\r' => 5,
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length
            .checked_add(escaped)
            .ok_or_else(|| Error::InvalidFormat("ODB settings output size overflow".to_string()))
    })
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
