//! Validated XLSB package ownership.
//!
//! This layer owns the OPC package boundary and the package-scoped BIFF12
//! readers and writers. The raw record kernel remains in [`crate::raw`],
//! while workbook and sheet handles live in [`crate::workbook`] and
//! [`crate::sheet`].

use std::io::{Read, Write};
use std::path::Path;

use litchi_opc::{OpcPackage, PackageWriter};

use crate::Workbook;
use crate::package::error::Result;
use crate::writer::{MutableWorksheet, WorkbookWriter};

#[path = "../host/cell.rs"]
pub(crate) mod cell;
#[path = "../host/cells_reader/mod.rs"]
#[allow(dead_code, unreachable_pub)]
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
#[path = "../host/pivot/mod.rs"]
pub mod pivot;
#[path = "../host/pivot_tables.rs"]
#[allow(dead_code, unreachable_pub)]
pub(crate) mod pivot_tables;
#[path = "../host/records.rs"]
#[allow(dead_code, unreachable_pub)]
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
#[allow(dead_code, unreachable_pub)]
pub(crate) mod utils;
#[path = "../host/vba_project.rs"]
pub mod vba_project;
#[path = "../host/walker.rs"]
pub(crate) mod walker;
#[path = "../host/web_extension_bindings.rs"]
pub mod web_extension_bindings;
#[path = "../host/xlsx/mod.rs"]
pub mod xlsx;

pub use cell::Cell;
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

    /// Open and validate an XLSB package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_opc(OpcPackage::open(path)?)
    }

    /// Read and validate an XLSB package from an owned byte vector.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_opc(OpcPackage::from_vec(bytes)?)
    }

    /// Read and validate an XLSB package from a borrowed byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_opc(OpcPackage::from_bytes(bytes)?)
    }

    /// Read and validate an XLSB package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_opc(OpcPackage::from_reader(reader)?)
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
