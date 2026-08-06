//! Immutable source snapshots for standalone Microsoft Graph packages.
//!
//! The package snapshot owns the exact compound-file artifact while retaining
//! the validated Workbook layout needed by the package transaction boundary.
//! Chart semantics are projected lazily so read-only callers do not pay for a
//! second chart allocation unless they inspect or edit the typed model.

use std::sync::Arc;

use crate::{Limits, Result, chart};

use super::validation::WorkbookLayout;
use super::{Package, Topology, codec, transaction, validation};

/// Exact, bounded source state for one standalone `[MS-OGRAPH]` package.
#[derive(Debug, Clone)]
pub struct Snapshot {
    pub(super) bytes: Arc<[u8]>,
    pub(super) topology: Topology,
    pub(super) workbook: WorkbookLayout,
    pub(super) limits: Limits,
}

impl Snapshot {
    /// Parses an owned standalone Graph package with default bounds.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parses a borrowed standalone Graph package while retaining an owned
    /// source copy.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes(bytes.to_vec())
    }

    /// Parses an owned package under explicit resource bounds.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        validation::check_limit("package bytes", bytes.len(), limits.max_package_bytes)?;
        let validated = validation::validate(&bytes, limits)?;
        let bytes = Arc::from(bytes.into_boxed_slice());
        Ok(Self {
            bytes,
            topology: validated.topology,
            workbook: validated.workbook,
            limits,
        })
    }

    pub(super) fn from_package(package: Package) -> Self {
        let (bytes, topology, workbook, limits) = package.into_snapshot_parts();
        Self {
            bytes: Arc::from(bytes.into_boxed_slice()),
            topology,
            workbook,
            limits,
        }
    }

    /// Exact compound-file source bytes.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Validated root-stream topology captured from the source.
    pub const fn topology(&self) -> Topology {
        self.topology
    }

    /// Resource limits retained for projections and subsequent edits.
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Parses the standalone Graph chart into its typed, inert semantic model.
    pub fn chart(&self) -> Result<chart::Chart> {
        codec::read_chart(self)
    }

    /// Starts a source-checked transaction over the typed Workbook chart.
    pub fn transaction(&self) -> Result<transaction::Transaction> {
        transaction::Transaction::new(self.clone())
    }

    /// Contextual alias for [`Self::transaction`].
    pub fn edit(&self) -> Result<transaction::Transaction> {
        self.transaction()
    }

    /// Recovers the owned source bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes.to_vec()
    }

    pub(super) fn source_bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub(super) fn source_arc(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }
}
