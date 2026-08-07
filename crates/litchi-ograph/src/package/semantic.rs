use crate::chart;
use crate::{Limits, Result};

use super::validation::WorkbookLayout;
use super::{codec, validation};

/// Validated root-stream topology of a standalone `OGraph` compound file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Topology {
    pub(super) workbook_bytes: u64,
    pub(super) comp_obj_bytes: Option<u64>,
    pub(super) ole_bytes: Option<u64>,
}

impl Topology {
    /// Declared byte size of the required `Workbook` stream.
    #[must_use]
    pub const fn workbook_bytes(self) -> u64 {
        self.workbook_bytes
    }

    /// Declared size of the optional `\u{1}CompObj` stream.
    #[must_use]
    pub const fn comp_obj_bytes(self) -> Option<u64> {
        self.comp_obj_bytes
    }

    /// Declared size of the optional `\u{1}Ole` stream.
    #[must_use]
    pub const fn ole_bytes(self) -> Option<u64> {
        self.ole_bytes
    }

    /// Number of allowed root streams present in the package.
    #[must_use]
    pub const fn stream_count(self) -> usize {
        1 + self.comp_obj_bytes.is_some() as usize + self.ole_bytes.is_some() as usize
    }
}

/// Borrowed, validated standalone `OGraph` package.
///
/// The compound-file bytes remain caller-owned. CFB stream extraction returns
/// an owned buffer because a logical stream can span non-contiguous sectors;
/// record traversal over that buffer is then zero-copy.
#[derive(Debug, Clone, Copy)]
pub struct PackageRef<'a> {
    bytes: &'a [u8],
    topology: Topology,
    workbook: WorkbookLayout,
    limits: Limits,
}

impl<'a> PackageRef<'a> {
    /// Validates borrowed bytes with conservative default limits.
    pub fn open(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Validates borrowed bytes with explicit resource limits.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        validation::check_limit("package bytes", bytes.len(), limits.max_package_bytes)?;
        let validated = validation::validate(bytes, limits)?;
        Ok(Self {
            bytes,
            topology: validated.topology,
            workbook: validated.workbook,
            limits,
        })
    }

    /// Original compound-file bytes.
    #[must_use]
    pub const fn as_bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// Validated stream topology.
    #[must_use]
    pub const fn topology(self) -> Topology {
        self.topology
    }

    /// Resource limits used to validate the package.
    #[must_use]
    pub const fn limits(self) -> Limits {
        self.limits
    }

    /// Reads the fragmented `Workbook` stream into one validated owned view.
    pub fn workbook(self) -> Result<Workbook> {
        let bytes = codec::read_workbook(self.bytes, self.workbook, self.limits)?;
        Ok(Workbook {
            bytes,
            layout: self.workbook,
            limits: self.limits,
        })
    }
}

/// Move-owned, validated standalone `OGraph` package.
#[derive(Debug)]
pub struct Package {
    bytes: Vec<u8>,
    topology: Topology,
    workbook: WorkbookLayout,
    limits: Limits,
}

impl Package {
    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates with explicit resource limits.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let validated = PackageRef::with_limits(&bytes, limits)?;
        let topology = validated.topology;
        let workbook = validated.workbook;
        let limits = validated.limits;
        Ok(Self {
            bytes,
            topology,
            workbook,
            limits,
        })
    }

    /// Borrows this package without revalidation or copying.
    #[must_use]
    pub fn as_ref(&self) -> PackageRef<'_> {
        PackageRef {
            bytes: &self.bytes,
            topology: self.topology,
            workbook: self.workbook,
            limits: self.limits,
        }
    }

    /// Converts this validated package into an exact source snapshot for
    /// contextual typed edits.
    #[must_use]
    pub fn snapshot(self) -> super::snapshot::Snapshot {
        super::snapshot::Snapshot::from_package(self)
    }

    /// Starts a source-checked transaction over the standalone Graph chart.
    pub fn edit(self) -> Result<super::transaction::Transaction> {
        self.snapshot().edit()
    }

    pub(super) fn into_snapshot_parts(self) -> (Vec<u8>, Topology, WorkbookLayout, Limits) {
        (self.bytes, self.topology, self.workbook, self.limits)
    }

    /// Validated stream topology.
    #[must_use]
    pub const fn topology(&self) -> Topology {
        self.topology
    }

    /// Reads the bounded standalone Workbook stream.
    pub fn workbook(&self) -> Result<Workbook> {
        self.as_ref().workbook()
    }

    /// Consumes the structurally validated package as an opaque attachment payload.
    ///
    /// This is a move: the compound-file allocation is neither cloned nor
    /// rebuilt.
    #[must_use]
    pub fn finish(self) -> Payload {
        Payload {
            bytes: self.bytes,
            topology: self.topology,
            workbook: self.workbook,
            limits: self.limits,
        }
    }
}

/// Opaque standalone `OGraph` bytes with validated CFB topology and BIFF framing.
///
/// This capability does not claim that the still-opaque chart records satisfy
/// the complete `[MS-OGRAPH]` chart grammar. Hosts that require that guarantee
/// must validate the typed chart model before attaching the payload.
#[derive(Debug)]
pub struct Payload {
    bytes: Vec<u8>,
    topology: Topology,
    workbook: WorkbookLayout,
    limits: Limits,
}

impl Payload {
    /// Takes ownership and validates bytes for attachment.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Package::open(bytes).map(Package::finish)
    }

    /// Takes ownership and validates bytes with explicit resource limits.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        Package::with_limits(bytes, limits).map(Package::finish)
    }

    /// Borrows the validated attachment bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Borrows the validated package view without copying.
    #[must_use]
    pub fn as_package(&self) -> PackageRef<'_> {
        PackageRef {
            bytes: &self.bytes,
            topology: self.topology,
            workbook: self.workbook,
            limits: self.limits,
        }
    }

    /// Recovers the original allocation without copying.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

/// Borrowed, validated standalone `OGraph` `Workbook` stream.
#[derive(Debug, Clone, Copy)]
pub struct WorkbookRef<'a> {
    bytes: &'a [u8],
    layout: WorkbookLayout,
    limits: Limits,
}

impl<'a> WorkbookRef<'a> {
    /// Validates the exact standalone globals-plus-chart topology.
    pub fn open(bytes: &'a [u8]) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Validates the standalone topology with explicit resource bounds.
    pub fn with_limits(bytes: &'a [u8], limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        validation::check_limit("Workbook bytes", bytes.len(), limits.max_workbook_bytes)?;
        let layout = validation::validate_workbook(bytes, limits)?;
        Ok(Self {
            bytes,
            layout,
            limits,
        })
    }

    /// Exact Workbook stream bytes.
    #[must_use]
    pub const fn as_bytes(self) -> &'a [u8] {
        self.bytes
    }

    /// The one standalone Microsoft Graph chart substream.
    #[must_use]
    pub fn chart(self) -> chart::Ref<'a> {
        let bytes = match self
            .bytes
            .get(self.layout.chart_start..self.layout.chart_end)
        {
            Some(bytes) => bytes,
            None => &[],
        };
        chart::Ref::from_validated(
            bytes,
            chart::Kind::Graph,
            self.layout.chart_start,
            self.limits,
        )
    }

    /// Resource limits under which the Workbook was validated.
    #[must_use]
    pub const fn limits(self) -> Limits {
        self.limits
    }
}

/// Move-owned, validated standalone `OGraph` `Workbook` stream.
#[derive(Debug)]
pub struct Workbook {
    bytes: Vec<u8>,
    layout: WorkbookLayout,
    limits: Limits,
}

impl Workbook {
    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::with_limits(bytes, Limits::default())
    }

    /// Takes ownership and validates under explicit resource bounds.
    pub fn with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let workbook = WorkbookRef::with_limits(&bytes, limits)?;
        let layout = workbook.layout;
        let limits = workbook.limits;
        Ok(Self {
            bytes,
            layout,
            limits,
        })
    }

    /// Borrow without copying or revalidation.
    #[must_use]
    pub fn as_ref(&self) -> WorkbookRef<'_> {
        WorkbookRef {
            bytes: &self.bytes,
            layout: self.layout,
            limits: self.limits,
        }
    }

    /// Exact Workbook stream bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// The one standalone Microsoft Graph chart substream.
    #[must_use]
    pub fn chart(&self) -> chart::Ref<'_> {
        self.as_ref().chart()
    }

    /// Recover the original Workbook allocation without copying.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.bytes
    }
}
