//! Source-checked, failure-atomic edits for the XLSB connections owner.

use litchi_opc::OpcPackage;

use super::model::{Connection, Connections, Parameter, UnknownRecord};
use super::{Result, package, validation};

/// An immutable workbook connections snapshot bound to the exact OPC owner
/// graph and BIFF12 source bytes.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    connections: Option<Connections>,
    source: package::SourceImage,
    unknown_records: Vec<UnknownRecord>,
}

impl Snapshot {
    /// Read and validate the complete workbook connections graph.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        let connections = package::validate_graph(package)?;
        let source = package::capture_source(package)?;
        let unknown_records = source
            .connection
            .as_ref()
            .map(|part| super::codec::parse_source(part.bytes.as_ref()))
            .transpose()?
            .map_or_else(Vec::new, |source| source.unknown_records);
        validation::unknown_records(&unknown_records)?;
        Ok(Self {
            connections,
            source,
            unknown_records,
        })
    }

    /// Alias for [`Snapshot::read`].
    pub fn load(package: &OpcPackage) -> Result<Self> {
        Self::read(package)
    }

    /// Borrow the typed connection catalog, when present.
    #[must_use]
    pub const fn connections(&self) -> Option<&Connections> {
        self.connections.as_ref()
    }

    /// Alias for the contextual connection catalog accessor.
    #[must_use]
    pub const fn catalog(&self) -> Option<&Connections> {
        self.connections()
    }

    /// Whether this snapshot has no connections owner.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.connections.is_none()
    }

    /// Exact source bytes of the BIFF12 connections part, when present.
    #[must_use]
    pub fn source_bytes(&self) -> Option<&[u8]> {
        self.source
            .connection
            .as_ref()
            .map(|part| part.bytes.as_ref().as_slice())
    }

    /// Borrow future/producer-specific records retained by this snapshot.
    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    /// A package snapshot is always source-bound.
    #[must_use]
    pub const fn is_source_bound(&self) -> bool {
        true
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

/// A detached, failure-atomic transaction over inert connection metadata.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<Connections>,
}

impl<'a> Transaction<'a> {
    /// Start a transaction against a validated package.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::read(target)?;
        Ok(Self {
            draft: before.connections.clone(),
            target,
            before,
        })
    }

    /// Borrow the source snapshot used for stale-source checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged connection catalog.
    #[must_use]
    pub const fn connections(&self) -> Option<&Connections> {
        self.draft.as_ref()
    }

    /// Replace or remove the complete connections owner.
    pub fn replace(&mut self, value: Option<Connections>) -> Result<bool> {
        if let Some(value) = &value {
            validation::connections(value)?;
        }
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    /// Edit one connection through a typed, failure-atomic closure.
    pub fn edit(
        &mut self,
        connection_id: u32,
        edit: impl FnOnce(&mut Connection) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| super::invalid("cannot edit an absent connections owner"))?;
        let connection = draft
            .connections
            .iter_mut()
            .find(|connection| connection.connection_id == connection_id)
            .ok_or_else(|| super::invalid("connection ID was not found"))?;
        edit(connection)?;
        if connection.connection_id != connection_id {
            return Err(super::invalid(
                "connection ID is immutable in a transaction",
            ));
        }
        validation::connections(&draft)?;
        if self.draft.as_ref() == Some(&draft) {
            return Ok(false);
        }
        self.draft = Some(draft);
        Ok(true)
    }

    /// Add a connection or replace one with the same stable identifier.
    pub fn set(&mut self, connection: Connection) -> Result<bool> {
        let mut draft = self.draft.clone().unwrap_or_default();
        if let Some(existing) = draft
            .connections
            .iter_mut()
            .find(|existing| existing.connection_id == connection.connection_id)
        {
            if *existing == connection {
                return Ok(false);
            }
            *existing = connection;
        } else {
            draft.connections.push(connection);
        }
        validation::connections(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Remove one connection by stable identifier.
    pub fn remove(&mut self, connection_id: u32) -> Result<Option<Connection>> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| super::invalid("cannot remove from an absent connections owner"))?;
        let Some(index) = draft
            .connections
            .iter()
            .position(|connection| connection.connection_id == connection_id)
        else {
            return Ok(None);
        };
        let removed = draft.connections.remove(index);
        if draft.connections.is_empty() {
            self.draft = None;
        } else {
            validation::connections(&draft)?;
            self.draft = Some(draft);
        }
        Ok(Some(removed))
    }

    /// Replace one parameter without contacting its provider.
    pub fn set_parameter(
        &mut self,
        connection_id: u32,
        index: usize,
        parameter: Parameter,
    ) -> Result<bool> {
        self.edit(connection_id, |connection| {
            let target = connection
                .parameters
                .get_mut(index)
                .ok_or_else(|| super::invalid("parameter index was not found"))?;
            *target = parameter;
            Ok(())
        })
    }

    /// Append one inert connection parameter.
    pub fn push_parameter(&mut self, connection_id: u32, parameter: Parameter) -> Result<bool> {
        self.edit(connection_id, |connection| {
            connection.parameters.push(parameter);
            Ok(())
        })
    }

    /// Remove one inert connection parameter.
    pub fn remove_parameter(
        &mut self,
        connection_id: u32,
        index: usize,
    ) -> Result<Option<Parameter>> {
        let mut removed = None;
        self.edit(connection_id, |connection| {
            if index >= connection.parameters.len() {
                return Err(super::invalid("parameter index was not found"));
            }
            removed = Some(connection.parameters.remove(index));
            Ok(())
        })?;
        Ok(removed)
    }

    /// Validate and publish the transaction as a reversible package patch.
    pub fn commit(self) -> Result<Commit> {
        if self.draft == self.before.connections {
            return Ok(Commit {
                snapshot: self.before.clone(),
                patch: Patch {
                    before: self.before.clone(),
                    after: self.before.clone(),
                },
                changed: false,
            });
        }

        let current = Snapshot::read(self.target)?;
        if !current.same_source(&self.before) {
            return Err(super::invalid("connections transaction source is stale"));
        }
        let mut candidate = self.target.clone();
        match &self.draft {
            Some(value) => {
                package::store_on_workbook(
                    &mut candidate,
                    &self.before.source.workbook_name,
                    value,
                )?;
            },
            None => {
                package::remove_from_workbook(&mut candidate, &self.before.source.workbook_name)?;
            },
        }
        let snapshot = Snapshot::read(&candidate)?;
        if snapshot.connections.as_ref() != self.draft.as_ref() {
            return Err(super::invalid(
                "connection publication changed the staged model",
            ));
        }
        *self.target = candidate;
        Ok(Commit {
            patch: Patch {
                before: self.before,
                after: snapshot.clone(),
            },
            snapshot,
            changed: true,
        })
    }
}

/// A successful source-checked package commit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the transaction changed the package.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the committed snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the commit into snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible package patch guarded by exact source graph metadata.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Whether applying this patch is source-identical.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Borrow the before-image snapshot.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the after-image snapshot.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Apply only to a package with the exact source graph used to create it.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        let current = Snapshot::read(target)?;
        if !current.same_source(&self.before) {
            return Err(super::invalid("connections patch source is stale"));
        }
        if self.is_empty() {
            return Ok(current);
        }
        let mut candidate = target.clone();
        package::restore_source(&mut candidate, &self.after.source)?;
        let snapshot = Snapshot::read(&candidate)?;
        if !snapshot.same_source(&self.after) {
            return Err(super::invalid(
                "connection patch publication changed its source",
            ));
        }
        *target = candidate;
        Ok(snapshot)
    }

    /// Alias for [`Patch::apply`].
    pub fn commit(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        self.apply(target)
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Read the source-checked connections owner.
pub fn read(package: &OpcPackage) -> Result<Snapshot> {
    Snapshot::read(package)
}

/// Apply a source-checked connections patch.
pub fn apply(target: &mut OpcPackage, patch: &Patch) -> Result<Snapshot> {
    patch.apply(target)
}
