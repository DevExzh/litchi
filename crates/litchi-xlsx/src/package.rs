//! Validated XLSX package ownership.
//!
//! [`Package`] is the physical-package boundary for the standalone XLSX
//! crate. It owns the OPC graph, delegates archive I/O to `litchi-opc`, and
//! validates the SpreadsheetML workbook graph before exposing a package
//! handle. Semantic reads and transactional edits live in [`crate::workbook`]
//! and [`crate::edit`].

/// Inert worksheet Printer Settings parts and references.
pub mod printer_settings;
/// Deterministic minimal package resources.
pub mod template;

use std::io::{Read, Write};
use std::path::Path;

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};

use crate::Workbook;
use crate::error::{Result, invalid};
use crate::writer;

/// A validated physical XLSX package.
///
/// The ZIP/container implementation remains below `litchi-opc`; this type
/// exposes only XLSX package operations and does not leak archive-specific
/// readers, writers, or errors.
#[derive(Debug, Clone)]
pub struct Package(OpcPackage);

impl Package {
    /// Create a deterministic minimal XLSX package with one visible worksheet.
    pub fn create() -> Result<Self> {
        Self::from_opc(build_minimal_package()?)
    }

    /// Open and validate an XLSX package from a filesystem path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_opc(OpcPackage::open(path)?)
    }

    /// Read and validate an XLSX package from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_opc(OpcPackage::from_vec(bytes)?)
    }

    /// Read and validate an XLSX package from a borrowed byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_opc(OpcPackage::from_bytes(bytes)?)
    }

    /// Read and validate an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_opc(OpcPackage::from_reader(reader)?)
    }

    /// Validate and adopt an already parsed OPC package.
    pub fn from_opc(package: OpcPackage) -> Result<Self> {
        // Run the complete workbook/relationship validation once at the
        // package boundary. The cloned graph shares immutable OPC payloads.
        Workbook::from_package(package.clone())?;
        Ok(Self(package))
    }

    /// Materialize an immutable workbook snapshot from this package.
    pub fn workbook(&self) -> Result<Workbook> {
        Workbook::from_package(self.0.clone())
    }

    /// Consume this package and materialize its immutable workbook snapshot.
    pub fn into_workbook(self) -> Result<Workbook> {
        Workbook::from_package(self.0)
    }

    /// Serialize the package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        writer::to_bytes(&self.0)
    }

    /// Stream the package into a sequential sink.
    pub fn write_to(&self, sink: impl Write) -> Result<()> {
        writer::write_to(&self.0, sink)
    }

    /// Atomically save the package to a filesystem path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        writer::save(&self.0, path)
    }
}

impl From<Package> for OpcPackage {
    fn from(package: Package) -> Self {
        package.0
    }
}

pub(crate) fn build_minimal_package() -> Result<OpcPackage> {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").map_err(invalid)?;
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").map_err(invalid)?;
    let styles_uri = PackURI::new("/xl/styles.xml").map_err(invalid)?;

    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
            r#"<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>"#,
            r#"</workbook>"#
        )
        .as_bytes()
        .to_vec(),
    );
    workbook.rels_mut().try_add_relationship(
        rt::WORKSHEET.to_owned(),
        "worksheets/sheet1.xml".to_owned(),
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    workbook.rels_mut().try_add_relationship(
        rt::STYLES.to_owned(),
        "styles.xml".to_owned(),
        "rId2".to_owned(),
        TargetMode::Internal,
    )?;
    package.try_add_part(Box::new(workbook))?;
    package.try_add_part(Box::new(BlobPart::new(
        worksheet_uri,
        ct::SML_WORKSHEET.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<dimension ref="A1"/><sheetData/></worksheet>"#
        )
        .as_bytes()
        .to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        styles_uri,
        ct::SML_STYLES.to_string(),
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
            r#"<fonts count="1"><font/></fonts>"#,
            r#"<fills count="2"><fill><patternFill patternType="none"/></fill>"#,
            r#"<fill><patternFill patternType="gray125"/></fill></fills>"#,
            r#"<borders count="1"><border/></borders>"#,
            r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
            r#"<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>"#,
            r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#,
            r#"</styleSheet>"#
        )
        .as_bytes()
        .to_vec(),
    )))?;
    package.rels_mut().try_add_relationship(
        rt::OFFICE_DOCUMENT.to_owned(),
        "xl/workbook.xml".to_owned(),
        "rId1".to_owned(),
        TargetMode::Internal,
    )?;
    Ok(package)
}

/// Inert analytical-model package resources.
pub mod xldm;
