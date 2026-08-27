//! Validated XLSB package ownership.
//!
//! This layer owns the OPC package boundary and the package-scoped BIFF12
//! readers and writers. The raw record kernel remains in [`crate::raw`],
//! while workbook and sheet handles live in [`crate::workbook`] and
//! [`crate::sheet`].

use std::io::{Read, Write};
use std::path::Path;

use litchi_core::sheet::WorkbookTrait;
use litchi_opc::{OpcPackage, PackageWriter, ReadLimits};

use crate::Workbook;
use crate::package::error::Result;
use crate::writer::{MutableWorksheet, WorkbookWriter};

#[path = "../host/cell.rs"]
pub(crate) mod cell;
#[path = "../host/cells_reader/mod.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod cells_reader;
#[path = "../host/chart_resources.rs"]
pub(crate) mod chart_resources;
#[path = "../host/chartsheet/mod.rs"]
pub mod chartsheet;
#[path = "../host/connections/mod.rs"]
pub mod connections;
#[path = "../host/data_validation.rs"]
pub mod data_validation;
#[path = "../host/drawing.rs"]
pub mod drawing;
#[path = "../host/drawing_image.rs"]
pub mod drawing_image;
#[path = "../host/drawing_write.rs"]
pub(crate) mod drawing_write;
#[path = "../host/error.rs"]
pub mod error;
pub(crate) mod external_link;
#[path = "../host/formula/mod.rs"]
pub mod formula;
#[path = "../host/frt.rs"]
pub(crate) mod frt;
#[path = "../host/named_ranges.rs"]
pub mod named_ranges;
pub(crate) mod owner_transaction;
#[path = "../host/pivot/mod.rs"]
pub mod pivot;
#[path = "../host/pivot_tables.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod pivot_tables;
#[path = "../host/records.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod records;
#[path = "../host/scenarios/mod.rs"]
pub mod scenarios;
#[path = "../host/shared_strings.rs"]
pub mod shared_strings;
#[path = "../host/sheet_view.rs"]
pub mod sheet_view;
#[path = "../host/styles/mod.rs"]
pub mod styles;
#[path = "../host/styles_table.rs"]
pub mod styles_table;
#[path = "../host/table/mod.rs"]
pub mod table;
#[path = "../host/template.rs"]
pub(crate) mod template;
#[path = "../host/utils.rs"]
#[allow(
    dead_code,
    unreachable_pub,
    reason = "internal host modules are shared by included package and writer implementations"
)]
pub(crate) mod utils;
#[cfg(feature = "vba-inspection")]
#[path = "../host/vba_project.rs"]
pub mod vba_project;
#[path = "../host/walker.rs"]
pub(crate) mod walker;
#[path = "../host/web_extension_bindings.rs"]
pub mod web_extension_bindings;

pub use cell::{Cell, FormulaOpacityReason, FormulaResolutionStatus};
pub use error::{Error as PackageError, Result as PackageResult};
pub use shared_strings::{
    PhoneticAlignment, PhoneticRun, PhoneticString, PhoneticType, SharedString, SharedStringRun,
};

/// A validated physical XLSB package.
///
/// ZIP/archive mechanics are delegated to `litchi-opc`; this type exposes
/// only the XLSB package boundary and validates the workbook graph before a
/// package is handed to callers.
#[derive(Debug, Clone)]
pub struct Package(OpcPackage);

impl Package {
    /// Create a deterministic workbook containing one visible worksheet.
    pub fn create() -> Result<Self> {
        let mut writer = WorkbookWriter::new();
        writer.add_worksheet(MutableWorksheet::new("Sheet1"));
        let mut bytes = std::io::Cursor::new(Vec::new());
        writer.save(&mut bytes)?;
        Self::from_bytes(bytes.into_inner())
    }

    /// Read workbook slicer caches as an immutable, source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the slicer owner or its dependency closure is invalid.
    pub fn slicer_caches(&self) -> Result<crate::slicer::Snapshot> {
        crate::slicer::transaction::read_caches(&self.0, &self.workbook_uri()?)
    }

    /// Read worksheet slicer views by checked zero-based worksheet position.
    ///
    /// `Ok(None)` means the worksheet selector did not resolve. An existing
    /// worksheet always returns a snapshot, including when it has no view part.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the slicer owner is invalid.
    pub fn slicer_views(&self, worksheet_index: usize) -> Result<Option<crate::slicer::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri(worksheet_index)? else {
            return Ok(None);
        };
        crate::slicer::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Read worksheet slicer views by exact worksheet name.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the slicer owner is invalid.
    pub fn slicer_views_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::slicer::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri_by_name(worksheet_name)? else {
            return Ok(None);
        };
        crate::slicer::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Apply a source-checked slicer patch and return a new package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch is stale or its exact readback fails.
    pub fn apply_slicer_patch(&self, patch: &crate::slicer::Patch) -> Result<Self> {
        let mut candidate = self.0.clone();
        crate::slicer::transaction::apply(&mut candidate, patch)?;
        Self::from_opc(candidate)
    }

    /// Read workbook timeline caches as an immutable, source-bound snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the timeline owner or its dependency closure is invalid.
    pub fn timeline_caches(&self) -> Result<crate::timeline::Snapshot> {
        crate::timeline::transaction::read_caches(&self.0, &self.workbook_uri()?)
    }

    /// Read worksheet timeline views by checked zero-based worksheet position.
    ///
    /// `Ok(None)` means the worksheet selector did not resolve. An existing
    /// worksheet always returns a snapshot, including when it has no view part.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the timeline owner is invalid.
    pub fn timeline_views(
        &self,
        worksheet_index: usize,
    ) -> Result<Option<crate::timeline::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri(worksheet_index)? else {
            return Ok(None);
        };
        crate::timeline::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Read worksheet timeline views by exact worksheet name.
    ///
    /// # Errors
    ///
    /// Returns an error when workbook resolution or the timeline owner is invalid.
    pub fn timeline_views_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::timeline::Snapshot>> {
        let Some(worksheet_uri) = self.worksheet_uri_by_name(worksheet_name)? else {
            return Ok(None);
        };
        crate::timeline::transaction::read_views(&self.0, &worksheet_uri).map(Some)
    }

    /// Apply a source-checked timeline patch and return a new package snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the patch is stale or its exact readback fails.
    pub fn apply_timeline_patch(&self, patch: &crate::timeline::Patch) -> Result<Self> {
        let mut candidate = self.0.clone();
        crate::timeline::transaction::apply(&mut candidate, patch)?;
        Self::from_opc(candidate)
    }

    /// Open and validate an XLSB package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open and validate an XLSB package from a filesystem path with explicit
    /// OPC resource limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::open_with_limits(path, limits)?)
    }

    /// Read and validate an XLSB package from an owned byte vector.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSB package from an owned byte vector with
    /// explicit OPC resource limits.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_vec_with_limits(bytes, limits)?)
    }

    /// Read and validate an XLSB package from a borrowed byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_slice_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSB package from a borrowed byte slice with
    /// explicit OPC resource limits.
    pub fn from_slice_with_limits(bytes: &[u8], limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_bytes_with_limits(bytes, limits)?)
    }

    /// Read and validate an XLSB package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Read and validate an XLSB package from a synchronous reader with
    /// explicit OPC resource limits.
    pub fn from_reader_with_limits(reader: impl Read, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_reader_with_limits(reader, limits)?)
    }

    /// Validate and adopt an already parsed OPC package.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        Workbook::from_opc_package(package.clone())?;
        Ok(Self(package))
    }

    /// Borrow the validated OPC package for lower-level package services.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.0
    }

    /// Materialize a workbook handle from this package.
    pub fn workbook(&self) -> Result<Workbook> {
        Workbook::from_opc_package(self.0.clone())
    }

    /// Consume this package and materialize its workbook handle.
    pub fn into_workbook(self) -> Result<Workbook> {
        Workbook::from_opc_package(self.0)
    }

    /// Consume this package and return its underlying OPC graph.
    pub fn into_opc(self) -> OpcPackage {
        self.0
    }

    /// Serialize the validated package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        Ok(PackageWriter::to_bytes(&self.0)?)
    }

    /// Stream the validated package into a sequential sink.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        Ok(PackageWriter::write_to_stream(writer, &self.0)?)
    }

    /// Atomically save the validated package to a filesystem path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        Ok(PackageWriter::write(path, &self.0)?)
    }

    fn worksheet_uri(&self, index: usize) -> Result<Option<litchi_opc::PackURI>> {
        let workbook = self.workbook()?;
        if index >= workbook.worksheet_names().len() {
            return Ok(None);
        }
        workbook.worksheet_uri(index).map(Some)
    }

    fn worksheet_uri_by_name(&self, name: &str) -> Result<Option<litchi_opc::PackURI>> {
        let workbook = self.workbook()?;
        let Some(index) = workbook
            .worksheet_names()
            .iter()
            .position(|worksheet| worksheet == name)
        else {
            return Ok(None);
        };
        workbook.worksheet_uri(index).map(Some)
    }

    fn workbook_uri(&self) -> Result<litchi_opc::PackURI> {
        Ok(self.0.main_document_part()?.partname().clone())
    }
}

impl From<Package> for OpcPackage {
    fn from(package: Package) -> Self {
        package.into_opc()
    }
}

impl From<OpcPackage> for Package {
    fn from(package: OpcPackage) -> Self {
        Self(package)
    }
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "package ingress fixtures use panic-on-failure extraction after constructing exact finite limits"
)]
mod tests {
    use std::io::Cursor;

    use super::*;

    #[test]
    fn ingress_honors_exact_read_limits() {
        let bytes = Package::create()
            .expect("create package")
            .to_bytes()
            .expect("serialize package");
        let input_bytes = u64::try_from(bytes.len()).expect("input length fits u64");
        let exact = ReadLimits::builder()
            .max_input_bytes(input_bytes)
            .expect("exact input limit")
            .build()
            .expect("valid exact limit");
        let over = ReadLimits::builder()
            .max_input_bytes(input_bytes - 1)
            .expect("smaller input limit")
            .build()
            .expect("valid smaller limit");
        let file = tempfile::NamedTempFile::new().expect("temporary workbook path");
        std::fs::write(file.path(), &bytes).expect("write workbook");

        assert!(Package::open(file.path()).is_ok());
        assert!(Package::from_bytes(bytes.clone()).is_ok());
        assert!(Package::from_slice(&bytes).is_ok());
        assert!(Package::from_reader(Cursor::new(bytes.clone())).is_ok());

        assert!(Package::open_with_limits(file.path(), exact).is_ok());
        assert!(Package::from_bytes_with_limits(bytes.clone(), exact).is_ok());
        assert!(Package::from_slice_with_limits(&bytes, exact).is_ok());
        assert!(Package::from_reader_with_limits(Cursor::new(bytes.clone()), exact).is_ok());

        assert!(Package::open_with_limits(file.path(), over).is_err());
        assert!(Package::from_bytes_with_limits(bytes.clone(), over).is_err());
        assert!(Package::from_slice_with_limits(&bytes, over).is_err());
        assert!(Package::from_reader_with_limits(Cursor::new(bytes), over).is_err());
    }
}
