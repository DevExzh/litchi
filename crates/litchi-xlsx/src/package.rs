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

use litchi_ooxml_common::custom::Host;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, ReadLimits, TargetMode};

use crate::Workbook;
use crate::custom::Props;
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
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open and validate an XLSX package from a filesystem path with explicit
    /// OPC resource limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::open_with_limits(path, limits)?)
    }

    /// Read and validate an XLSX package from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSX package from owned bytes with explicit OPC
    /// resource limits.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_vec_with_limits(bytes, limits)?)
    }

    /// Read and validate an XLSX package from a borrowed byte slice.
    pub fn from_slice(bytes: &[u8]) -> Result<Self> {
        Self::from_slice_with_limits(bytes, ReadLimits::default())
    }

    /// Read and validate an XLSX package from a borrowed byte slice with
    /// explicit OPC resource limits.
    pub fn from_slice_with_limits(bytes: &[u8], limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_bytes_with_limits(bytes, limits)?)
    }

    /// Read and validate an XLSX package from a synchronous reader.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Read and validate an XLSX package from a synchronous reader with
    /// explicit OPC resource limits.
    pub fn from_reader_with_limits(reader: impl Read, limits: ReadLimits) -> Result<Self> {
        Self::from_opc(OpcPackage::from_reader_with_limits(reader, limits)?)
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

    /// Read the inert typed package-level custom document properties.
    ///
    /// An absent custom-properties relationship and part produce the shared
    /// empty [`Props`] value. The custom-properties package graph follows
    /// MS-OI29500 section 3.11.
    ///
    /// Values are metadata only: this method does not evaluate formulas or
    /// execute macros, controls, VBA, links, external code, or payloads.
    pub fn custom_props(&self) -> Result<Props> {
        Ok(Props::read_for(&self.0, Host::Excel)?)
    }

    /// Atomically publish inert typed package-level custom document properties.
    ///
    /// An empty collection removes the custom-properties graph. The change is
    /// clone-staged and only replaces this facade's OPC package after the
    /// shared OOXML owner has validated and serialized the complete graph.
    /// Changed metadata invalidates existing package signatures.
    pub fn put_custom_props(&mut self, props: Props) -> Result<()> {
        self.stage(|package| {
            props.write_for(package, Host::Excel)?;
            Ok(())
        })
    }

    /// Atomically remove the package-level custom document properties.
    ///
    /// This idempotent operation rejects corrupt, external, or duplicate
    /// graphs instead of treating them as absent. The custom-properties graph
    /// is defined by MS-OI29500 section 3.11.
    pub fn remove_custom_props(&mut self) -> Result<()> {
        self.stage(|package| {
            Props::read_for(package, Host::Excel)?;
            Props::new().write_for(package, Host::Excel)?;
            Ok(())
        })
    }

    /// Serialize the package into owned bytes.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        writer::to_bytes(&self.0)
    }

    /// Stream the package into a sequential sink.
    pub fn write_to(&self, sink: impl Write) -> Result<()> {
        writer::write_to(&self.0, sink)
    }

    /// Start a clone-staged transaction for worksheet slicers.
    ///
    /// Dropping the returned transaction rolls back; `commit` publishes the
    /// validated feature graph without rebuilding unrelated workbook parts.
    pub fn edit_slicers(&mut self) -> Result<crate::slicer::Transaction<'_>> {
        crate::slicer::Transaction::new(&mut self.0)
    }

    /// Start a clone-staged transaction for worksheet timelines.
    ///
    /// Timeline cache references are workbook-owned, so the workbook part
    /// identity is captured once at transaction start and is not exposed by
    /// the ordinary package facade.
    pub fn edit_timelines(&mut self) -> Result<crate::timeline::Transaction<'_>> {
        let workbook = self.0.main_document_part()?.partname().clone();
        crate::timeline::Transaction::new(&mut self.0, &workbook)
    }

    /// Load the inert Office Add-in task-pane graph, when present.
    pub fn task_panes(&self) -> Result<Option<litchi_ooxml_common::web::Panes>> {
        crate::task_panes::load(&self.0)
    }

    /// Start a clone-staged transaction for the package-level task-pane graph.
    ///
    /// Existing XML conformance is retained automatically. New graphs use
    /// Transitional SpreadsheetML relationships unless the explicit
    /// [`Self::edit_task_panes_with`] entry point is selected.
    pub fn edit_task_panes(&mut self) -> Result<crate::task_panes::Transaction<'_>> {
        let conformance = crate::task_panes::existing_conformance(&self.0)?;
        crate::task_panes::Transaction::new(&mut self.0, conformance)
    }

    /// Start a task-pane transaction with an explicit XML relationship
    /// conformance. The staged graph is published only by
    /// [`crate::task_panes::Transaction::commit`].
    pub fn edit_task_panes_with(
        &mut self,
        conformance: litchi_ooxml_common::web::Conformance,
    ) -> Result<crate::task_panes::Transaction<'_>> {
        crate::task_panes::Transaction::new(&mut self.0, conformance)
    }

    /// Read one worksheet's inert smart-tag annotations by semantic selector.
    pub fn smart_tags<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Option<crate::smart_tags::Collection>> {
        let workbook = self.workbook()?;
        let worksheet = workbook
            .sheet(sheet)?
            .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
        worksheet.smart_tags()
    }

    /// Start an atomic smart-tag transaction for a semantic worksheet.
    pub fn edit_smart_tags<'a>(
        &mut self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<crate::smart_tags::Transaction<'_>> {
        let worksheet = {
            let workbook = self.workbook()?;
            let worksheet = workbook
                .sheet(sheet)?
                .ok_or_else(|| invalid("worksheet selector did not match a sheet"))?;
            if worksheet.kind() != crate::workbook::WorksheetKind::Worksheet {
                return Err(crate::Error::NotWorksheet {
                    sheet: worksheet.name().to_owned(),
                });
            }
            worksheet.part_uri().clone()
        };
        crate::smart_tags::Transaction::new(&mut self.0, worksheet)
    }

    /// Atomically save the package to a filesystem path.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        writer::save(&self.0, path)
    }

    fn stage<T>(&mut self, mutation: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut staged = self.0.clone();
        let value = mutation(&mut staged)?;
        self.0 = staged;
        Ok(value)
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
