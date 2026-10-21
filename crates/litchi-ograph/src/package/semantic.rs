use std::io::Cursor;
use std::sync::{Arc, Mutex, OnceLock};

use litchi_biff::{Encoder, Records};
use litchi_cfb::OleWriter;

use crate::chart;
use crate::{Limits, OpenOptions, Result};

use super::validation::WorkbookLayout;
use super::{codec, validation};

/// Validated root-stream topology of a standalone `OGraph` compound file.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::struct_field_names,
    reason = "the suffix distinguishes declared byte sizes from decoded stream values"
)]
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
        Self::open_with_options(bytes, OpenOptions::default())
    }

    /// Validates borrowed bytes with explicit resource limits.
    // Retain the caller's policy at this boundary so default scalar checks can
    // fold away. Structural admission remains in the shared codec.
    #[inline(always)]
    pub fn open_with_options(bytes: &'a [u8], options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value();
        let limits = limits.validate()?;
        validation::check_limit(
            crate::LimitKind::PackageBytes,
            bytes.len(),
            limits.max_package_bytes,
        )?;
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

/// Retains the caller's Vec until a shared borrow is requested. Promotion
/// moves only its header; existing bytes and capacity never change. The empty
/// Vec left in `owned` has no allocation. OnceLock serializes promotion, while
/// Mutex permits the move through &self without exposing a mutable byte slice.
#[derive(Debug)]
pub(super) struct Bytes {
    owned: Mutex<Vec<u8>>,
    shared: OnceLock<Arc<Vec<u8>>>,
}

impl Bytes {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            owned: Mutex::new(bytes),
            shared: OnceLock::new(),
        }
    }

    fn from_shared(bytes: Arc<Vec<u8>>) -> Self {
        Self {
            owned: Mutex::new(Vec::new()),
            shared: OnceLock::from(bytes),
        }
    }

    pub(super) fn shared(&self) -> &Arc<Vec<u8>> {
        self.shared.get_or_init(|| {
            let mut bytes = self
                .owned
                .lock()
                .unwrap_or_else(std::sync::PoisonError::into_inner);
            Arc::new(std::mem::take(&mut *bytes))
        })
    }

    #[inline]
    fn into_vec(self) -> Vec<u8> {
        match self.shared.into_inner() {
            Some(bytes) => Arc::unwrap_or_clone(bytes),
            None => self
                .owned
                .into_inner()
                .unwrap_or_else(std::sync::PoisonError::into_inner),
        }
    }
}

impl Clone for Bytes {
    fn clone(&self) -> Self {
        Self::from_shared(Arc::clone(self.shared()))
    }
}

impl std::ops::Deref for Bytes {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.shared()
    }
}

/// Immutable, cheaply shared standalone `OGraph` package.
#[derive(Debug, Clone)]
pub struct Package {
    pub(super) bytes: Bytes,
    pub(super) topology: Topology,
    pub(super) workbook: WorkbookLayout,
    pub(super) limits: Limits,
}

impl Package {
    /// Builds a standalone Graph package under explicit resource bounds.
    pub fn from_chart(chart: chart::Model, options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value().validate()?;
        let chart = chart.finish_graph_authoring()?;
        let chart_stream = chart.encode_with(limits)?.into_bytes();
        let mut workbook = Encoder::new(limits.biff)?;
        workbook.push(validation::BOF, &bof(validation::GLOBALS))?;
        workbook.push(validation::EOF, &[])?;
        for record in Records::new(&chart_stream, limits.biff)? {
            workbook.push_ref(record?)?;
        }
        let workbook = workbook.finish();
        validation::check_limit(
            crate::LimitKind::WorkbookBytes,
            workbook.len(),
            limits.max_workbook_bytes,
        )?;

        let mut writer = OleWriter::new();
        writer.create_stream(&[codec::WORKBOOK], &workbook)?;
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output)?;
        let output = output.into_inner();
        validation::check_limit(
            crate::LimitKind::PackageBytes,
            output.len(),
            limits.max_package_bytes,
        )?;
        Self::open_with_options(output, OpenOptions::new().limits(limits))
    }

    /// Takes ownership and validates without copying the input allocation.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_with_options(bytes, OpenOptions::default())
    }

    /// Takes ownership and validates with explicit resource limits.
    pub fn open_with_options(bytes: Vec<u8>, options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value();
        let validated = PackageRef::open_with_options(&bytes, OpenOptions::new().limits(limits))?;
        let topology = validated.topology;
        let workbook = validated.workbook;
        let limits = validated.limits;
        Ok(Self {
            bytes: Bytes::new(bytes),
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

    /// Exact compound-file bytes, including opaque streams.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Starts an edit without reading or copying package content.
    ///
    /// Chart semantics are validated on the first staged operation or commit.
    pub fn edit(&self) -> super::Edit {
        super::Edit::new(self.clone())
    }

    /// Replays a patch against any package with identical recorded base bytes.
    pub fn apply(&self, patch: &super::Patch) -> Result<super::Commit> {
        patch.apply_to(self)
    }

    pub(super) fn from_shared(bytes: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        let view = PackageRef::open_with_options(&bytes, OpenOptions::new().limits(limits))?;
        let topology = view.topology;
        let workbook = view.workbook;
        let limits = view.limits;
        Ok(Self {
            bytes: Bytes::from_shared(bytes),
            topology,
            workbook,
            limits,
        })
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
    /// Recovers the original allocation when this package is uniquely owned;
    /// otherwise copies its shared bytes without rebuilding the compound file.
    #[must_use]
    #[inline]
    pub fn finish(self) -> Payload {
        Payload {
            bytes: self.bytes.into_vec(),
            topology: self.topology,
            workbook: self.workbook,
            limits: self.limits,
        }
    }
}

fn bof(doc_type: u16) -> [u8; validation::BOF_BYTES] {
    let mut payload = [0; validation::BOF_BYTES];
    payload[0..2].copy_from_slice(&validation::OGRAPH_VERSION.to_le_bytes());
    payload[2..4].copy_from_slice(&doc_type.to_le_bytes());
    payload[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
    payload[6..8].copy_from_slice(&validation::OGRAPH_YEAR_1997.to_le_bytes());
    payload[8..12]
        .copy_from_slice(&(validation::REQUIRED_PLATFORM_FLAGS | (6 << 14)).to_le_bytes());
    payload[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
    payload
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
    pub fn open_with_options(bytes: Vec<u8>, options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value();
        Package::open_with_options(bytes, OpenOptions::new().limits(limits)).map(Package::finish)
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
        Self::open_with_options(bytes, OpenOptions::default())
    }

    /// Validates the standalone topology with explicit resource bounds.
    pub fn open_with_options(bytes: &'a [u8], options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value();
        let limits = limits.validate()?;
        validation::check_limit(
            crate::LimitKind::WorkbookBytes,
            bytes.len(),
            limits.max_workbook_bytes,
        )?;
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
        Self::open_with_options(bytes, OpenOptions::default())
    }

    /// Takes ownership and validates under explicit resource bounds.
    pub fn open_with_options(bytes: Vec<u8>, options: OpenOptions) -> Result<Self> {
        let limits = options.limits_value();
        let workbook = WorkbookRef::open_with_options(&bytes, OpenOptions::new().limits(limits))?;
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
