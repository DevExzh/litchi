//! Immutable edits of the XLSB External Data Connections owner (record
//! 0360b).
//!
//! [`Package::edit_connections`] starts an [`Edit`] over the workbook's
//! inert connection catalog; its verbs stage a replacement, one connection
//! or one parameter, each validated as a complete catalog. [`Edit::commit`]
//! stores the staged catalog (or removes the owner), reads it back, and
//! returns a [`Commit`] with the new package, the reversible [`Patch`] and
//! the [`Diagnostics`]; [`Package::apply_connections`] replays a patch onto
//! an equal package. Connection strings, commands, URLs, paths and
//! credential metadata are stored verbatim and never resolved or executed.
//! A catalog equal to the stored one publishes nothing, and a changed edit
//! of a signed workbook is refused with [`Error::Signed`].

use litchi_opc::OpcPackage;

use super::model::{Connection, Connections, Parameter, UnknownRecord};
use super::{Result, package, validation};
use crate::part_edit::{Staged, family};

family!("connections");

/// An immutable workbook connections snapshot bound to the exact OPC owner
/// graph and BIFF12 source bytes.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    connections: Option<Connections>,
    source: package::SourceImage,
    unknown_records: Vec<UnknownRecord>,
}

impl Snapshot {
    /// Read and validate the complete workbook connections graph under the
    /// owner's raw profile.
    pub(crate) fn read(package: &OpcPackage, raw_limits: crate::raw::Limits) -> Result<Self> {
        let connections = package::load(package, raw_limits)?;
        let source = package::capture_source(package)?;
        let unknown_records = source
            .connection
            .as_ref()
            .map(|part| super::codec::parse_source(part.bytes.as_ref(), raw_limits))
            .transpose()?
            .map_or_else(Vec::new, |source| source.unknown_records);
        validation::unknown_records(&unknown_records)?;
        Ok(Self {
            connections,
            source,
            unknown_records,
        })
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
}

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Connections of one catalog.
    Connections,
    /// Bytes of the External Data Connections part.
    PartBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Connections => "External Data Connections",
            Self::PartBytes => "External Data Connections bytes",
        })
    }
}

/// A staged connections edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    before: Snapshot,
    draft: Option<Connections>,
}

impl crate::Package {
    /// The source-bound External Data Connections snapshot of this package.
    ///
    /// # Errors
    ///
    /// Returns the connections graph and decoding errors.
    pub fn connections(&self) -> Result<Snapshot> {
        Snapshot::read(self.opc_package(), self.raw_limits_value())
    }

    /// Start a connections edit. Starting one copies no part payload.
    ///
    /// # Errors
    ///
    /// Returns the connections graph and decoding errors.
    pub fn edit_connections(&self) -> Result<Edit> {
        let before = Snapshot::read(self.opc_package(), self.raw_limits_value())?;
        Ok(Edit {
            staged: Staged::of_package(self)?,
            draft: before.connections.clone(),
            before,
        })
    }

    /// Apply a connections patch to a new package; this one is unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_connections(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

impl Edit {
    /// The source snapshot the edit started from.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The staged connection catalog.
    #[must_use]
    pub const fn connections(&self) -> Option<&Connections> {
        self.draft.as_ref()
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.draft = self.before.connections.clone();
    }

    /// Replace or remove the complete connections owner. Returns whether the
    /// staged catalog changed.
    ///
    /// # Errors
    ///
    /// Returns the catalog validation errors.
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
    ///
    /// # Errors
    ///
    /// Returns the closure's error, an absent owner or connection, a changed
    /// connection ID and the catalog validation errors.
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
    ///
    /// # Errors
    ///
    /// Returns the catalog validation errors.
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

    /// Remove one connection by stable identifier; removing the last one
    /// removes the owner.
    ///
    /// # Errors
    ///
    /// Returns an absent owner and the catalog validation errors.
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
    ///
    /// # Errors
    ///
    /// The errors of [`Self::edit`] and a missing parameter index.
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
    ///
    /// # Errors
    ///
    /// The errors of [`Self::edit`].
    pub fn push_parameter(&mut self, connection_id: u32, parameter: Parameter) -> Result<bool> {
        self.edit(connection_id, |connection| {
            connection.parameters.push(parameter);
            Ok(())
        })
    }

    /// Remove one inert connection parameter.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::edit`] and a missing parameter index.
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

    /// Store the staged catalog, read it back and publish a new package. An
    /// unchanged catalog returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed edit of a signed workbook, the
    /// connections graph errors, a readback that differs from the staged
    /// catalog, and the workbook validation errors of the result.
    pub fn commit(mut self) -> Result<Commit> {
        if self.draft == self.before.connections {
            self.staged.reset();
            return Ok(Commit::new(self.staged.commit()?));
        }
        crate::signature_policy::ensure_unsigned(self.staged.opc())?;
        let workbook = self.before.source.workbook_name.clone();
        let draft = self.draft.clone();
        let raw_limits = self.staged.raw_limits();
        self.staged.update(|candidate| {
            let stored = match &draft {
                Some(value) => Some(package::store_on_workbook(
                    candidate, &workbook, value, raw_limits,
                )?),
                None => {
                    package::remove_from_workbook(candidate, &workbook)?;
                    None
                },
            };
            let snapshot = Snapshot::read(candidate, raw_limits)?;
            if snapshot.connections != stored {
                return Err(super::invalid(
                    "connection publication changed the staged model",
                ));
            }
            Ok(())
        })?;
        Ok(Commit::new(self.staged.commit()?))
    }
}
