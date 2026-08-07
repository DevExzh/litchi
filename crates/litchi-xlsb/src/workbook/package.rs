//! OPC/package integration for the typed XLSB workbook.

use super::model::Workbook;
use crate::calc::Props;
use crate::cell_watches;
use crate::package::error::Result;
use crate::package::formula::{Context, View, excel_name_eq, table::Definition as TableDefinition};
use crate::package::styles_table::StylesTable;
use crate::package::vba_project::{
    VbaProject, discover_vba_project, remove_vba_project as clear_workbook_vba,
    store_vba_project as store_workbook_vba_project,
};
use crate::package::web_extension_bindings::PackageAppRefs;
use crate::raw::Records;
use litchi_ooxml_common::embedded;
use litchi_ooxml_common::ribbon;
use litchi_ooxml_common::web;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type, relationship_type};
use std::collections::HashMap;
use std::io::{Read, Seek, Write};
use std::sync::Arc;

/// Chart sheet relationship types documented by MS-XLSB 2.1.7.7.
const CHART_SHEET_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
];

impl Workbook {
    /// Read the typed cell-watch and worksheet phonetic snapshot selected by
    /// zero-based worksheet index.
    pub fn cell_watches(&self, worksheet_index: usize) -> Result<cell_watches::Snapshot> {
        let uri = self.worksheet_uri(worksheet_index)?;
        cell_watches::workbook::read(&self.package, &uri)
    }

    /// Start a detached cell-watch edit for one worksheet.
    pub fn edit_cell_watches(&self, worksheet_index: usize) -> Result<cell_watches::Edit> {
        Ok(self.cell_watches(worksheet_index)?.edit())
    }

    /// Apply a source-checked cell-watch commit atomically to one worksheet.
    pub fn apply_cell_watches(
        &mut self,
        worksheet_index: usize,
        commit: &cell_watches::Commit,
    ) -> Result<cell_watches::Snapshot> {
        let uri = self.worksheet_uri(worksheet_index)?;
        cell_watches::workbook::apply(&mut self.package, &uri, commit)
    }

    /// Read the typed cell-watch and phonetic snapshot selected by worksheet
    /// name.
    pub fn cell_watches_by_name(&self, worksheet_name: &str) -> Result<cell_watches::Snapshot> {
        let index = self.worksheet_index(worksheet_name)?;
        self.cell_watches(index)
    }

    /// Load inert persisted Office Add-in task panes.
    pub fn task_panes(&self) -> Result<Option<web::Panes>> {
        Ok(web::load(&self.package)?)
    }

    /// Store task panes after validating every binary worksheet `appRef`.
    pub fn put_task_panes(
        &mut self,
        panes: web::Panes,
        conformance: web::Conformance,
    ) -> Result<&mut Self> {
        self.validate_task_pane_bindings(&panes)?;
        web::put(&mut self.package, panes, conformance)?;
        Ok(self)
    }

    /// Remove task panes only when no binary worksheet binding would dangle.
    pub fn remove_task_panes(&mut self) -> Result<bool> {
        if !self
            .package
            .rels()
            .iter()
            .any(|relationship| relationship.reltype() == web::raw::TASK_PANES_RELATIONSHIP)
        {
            return Ok(false);
        }
        self.validate_task_pane_bindings(&web::Panes::new())?;
        Ok(web::remove(&mut self.package)?)
    }

    fn validate_task_pane_bindings(&self, panes: &web::Panes) -> Result<()> {
        let package_refs = PackageAppRefs::new(
            panes
                .iter()
                .flat_map(|pane| pane.add_in().bindings().iter()),
        )?;
        for worksheet in &self.worksheets {
            package_refs.validate(worksheet.web_extension_bindings())?;
        }
        Ok(())
    }

    /// Read the bounded, inert package-level Ribbon customizations.
    pub fn ribbon(&self) -> Result<ribbon::Set<'_>> {
        Ok(ribbon::load(&self.package)?)
    }

    /// Create or replace one Ribbon customization family.
    pub fn put_ribbon(&mut self, version: ribbon::Version, xml: Vec<u8>) -> Result<&mut Self> {
        ribbon::put(&mut self.package, version, xml)?;
        Ok(self)
    }

    /// Remove one Ribbon relationship family and its unreferenced part.
    pub fn remove_ribbon(&mut self, family: ribbon::Family) -> Result<bool> {
        Ok(ribbon::remove(&mut self.package, family)?)
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan_with`] with [`Self::opc_package`] when a lower
    /// layer needs explicitly tuned limits.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.package)?)
    }

    /// Get the underlying OPC package.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.package
    }

    /// Get mutable OPC access for XLSB-internal package adapters.
    ///
    /// Public callers must use [`Self::edit_opc`], which stages a structural
    /// candidate and reparses the complete XLSB host before publication.
    #[allow(dead_code)]
    pub(crate) fn opc_package_mut(&mut self) -> &mut OpcPackage {
        self.package.unsign();
        &mut self.package
    }

    /// Transactionally edit the current XLSB OPC graph.
    ///
    /// The closure receives a cloned candidate package. Returning an error,
    /// producing a package whose main relationship or XLSB graph is invalid,
    /// or unwinding leaves this workbook unchanged. A successful edit drops
    /// package signatures, reparses workbook-owned state, and revalidates the
    /// inert VBA and External Data Connections relationship graphs before
    /// publication.
    pub fn edit_opc<T>(&mut self, edit: impl FnOnce(&mut OpcPackage) -> Result<T>) -> Result<T> {
        let mut candidate = self.package.clone();
        candidate.unsign();
        let value = edit(&mut candidate)?;

        Self::validate_edit_candidate(&candidate)?;
        let validated = Self::from_opc_package(candidate)?;
        *self = validated;
        Ok(value)
    }

    fn validate_edit_candidate(package: &OpcPackage) -> Result<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook = package.get_part(&workbook_uri)?;
        if workbook.content_type() != content_type::XLSB_BIN {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "XLSB workbook content type".to_string(),
                val: format!(
                    "expected '{}', found '{}'",
                    content_type::XLSB_BIN,
                    workbook.content_type()
                ),
            });
        }

        let main = package.main_document_part()?;
        if main.partname() != workbook.partname() || main.content_type() != content_type::XLSB_BIN {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "XLSB main workbook relationship".to_string(),
                val: format!(
                    "expected '{}' to be the binary workbook main part",
                    workbook_uri.as_str()
                ),
            });
        }

        // These readers validate relationship cardinality, target mode,
        // target content type, and orphan/inbound graph invariants without
        // parsing or executing opaque payload bytes.
        discover_vba_project(package, main)?;
        Ok(())
    }

    /// Return whether this workbook contains package signatures.
    #[must_use]
    #[inline]
    pub fn is_signed(&self) -> bool {
        self.package.is_signed()
    }

    /// Verify package signatures with the safe strict policy.
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.package.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.package.signatures_with(policy)
    }

    /// Add a signature while preserving every existing valid signature.
    pub fn sign(
        &mut self,
        signer: &litchi_sign::Signer,
    ) -> litchi_opc::sign::Result<litchi_opc::PackURI> {
        self.package.sign(signer)
    }

    /// Add a signature with explicit authoring resource bounds.
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<litchi_opc::PackURI> {
        self.package.sign_with(signer, limits)
    }

    /// Atomically replace all signatures with one signature.
    pub fn resign(
        &mut self,
        signer: &litchi_sign::Signer,
    ) -> litchi_opc::sign::Result<litchi_opc::PackURI> {
        self.package.resign(signer)
    }

    /// Atomically replace signatures with explicit authoring resource bounds.
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> litchi_opc::sign::Result<litchi_opc::PackURI> {
        self.package.resign_with(signer, limits)
    }

    /// Remove all package signatures.
    pub fn unsign(&mut self) -> &mut Self {
        self.package.unsign();
        self
    }

    /// Discover the attached MS-XLSB VBA project and declared signatures.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, verify, or execute VBA project or
    /// signature bytes.
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let workbook = self.package.main_document_part()?;
        discover_vba_project(&self.package, workbook)
    }

    /// Attach a cache-free, inert MS-OVBA project to this binary workbook.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> Result<VbaProject> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Attach a cache-free project with explicit resource limits.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> Result<VbaProject> {
        self.put_vba(project.finish(limits)?)
    }

    /// Attach a prevalidated `vbaProject.bin` without executing it.
    ///
    /// Any existing legacy or Agile project signature is removed because
    /// replacing the signed project bytes invalidates it.
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> Result<VbaProject> {
        let source = self.package.main_document_part()?.partname().clone();
        store_workbook_vba_project(&mut self.package, &source, payload)
    }

    /// Remove the VBA project and all declared project-signature parts.
    pub fn clear_vba(&mut self) -> Result<bool> {
        let source = self.package.main_document_part()?.partname().clone();
        clear_workbook_vba(&mut self.package, &source)
    }

    /// The typed External Data Connections part, when the workbook declares
    /// one (MS-XLSB 2.1.7.24).
    ///
    /// These are inert data snapshots: connection strings, commands, URLs,
    /// file paths, and credential metadata are stored verbatim and are never
    /// resolved, contacted, refreshed, or executed.
    pub fn connections(&self) -> Option<&crate::package::connections::Connections> {
        self.connections.as_ref()
    }

    /// Atomically add or replace the inert External Data Connections part.
    ///
    /// Existing package content is preserved. Connection strings, commands,
    /// URLs, paths, and credential metadata are never resolved or executed.
    pub fn set_connections(
        &mut self,
        connections: crate::package::connections::Connections,
    ) -> Result<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let canonical = crate::package::connections::package::store_on_workbook(
            &mut self.package,
            &workbook_uri,
            &connections,
        )?;
        self.connections = Some(canonical);
        Ok(())
    }

    /// Remove the External Data Connections relationship and part.
    ///
    /// Returns `true` when a graph was removed.
    pub fn remove_connections(&mut self) -> Result<bool> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let removed = crate::package::connections::package::remove_from_workbook(
            &mut self.package,
            &workbook_uri,
        )?;
        if removed {
            self.connections = None;
        }
        Ok(removed)
    }
    fn load_sheet_drawing(
        &self,
        sheet_index: usize,
        drawing_part: &dyn litchi_opc::part::Part,
    ) -> Result<crate::package::drawing::SheetDrawing> {
        use crate::package::drawing::{EmbeddedChart, EmbeddedImage, Object, SheetDrawing};
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).map_err(|error| {
            crate::package::error::Error::Encoding(format!("Drawings part is not UTF-8: {error}"))
        })?;
        let shapes = crate::shapes::parse_drawing_shapes(drawing_xml)?.unwrap_or_default();
        let drawing = crate::package::drawing::parse_drawing_part(drawing_part.blob())?;
        let mut charts = Vec::new();
        let mut images = Vec::new();
        let mut image_bytes = 0usize;
        let mut image_cache = HashMap::new();
        for anchor in &drawing.anchors {
            if let Object::Picture {
                non_visual,
                embed_rel_id: Some(rel_id),
            } = &anchor.object
            {
                let relationship = drawing_part.rels().get(rel_id).ok_or_else(|| {
                    crate::package::error::Error::Unrecognized {
                        typ: "Drawings part".to_string(),
                        val: format!(
                            "picture {:?} relationship {rel_id:?} is missing",
                            non_visual.name
                        ),
                    }
                })?;
                if matches!(
                    relationship.reltype(),
                    relationship_type::IMAGE | relationship_type::STRICT_IMAGE
                ) {
                    if relationship.is_external() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "Drawings part".to_string(),
                            val: format!("image relationship {rel_id:?} is external"),
                        });
                    }
                    let image_uri = relationship.target_partname()?;
                    let image_part = self.package.get_part(&image_uri)?;
                    let Some(format) =
                        crate::package::drawing_image::ImageFormat::from_content_type(
                            image_part.content_type(),
                        )
                    else {
                        continue;
                    };
                    if images.len() >= crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGES {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGES,
                            found: images.len() + 1,
                        });
                    }
                    let data = if let Some(data) = image_cache.get(&image_uri) {
                        Arc::clone(data)
                    } else {
                        format.validate_payload(image_part.blob())?;
                        image_bytes = image_bytes
                            .checked_add(image_part.blob().len())
                            .ok_or(crate::package::error::Error::InvalidLength {
                            expected:
                                crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                            found: usize::MAX,
                        })?;
                        if image_bytes
                            > crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES
                        {
                            return Err(crate::package::error::Error::InvalidLength {
                                expected:
                                    crate::package::drawing_image::MAX_XLSB_WORKSHEET_IMAGE_TOTAL_BYTES,
                                found: image_bytes,
                            });
                        }
                        let data = Arc::<[u8]>::from(image_part.blob());
                        image_cache.insert(image_uri, Arc::clone(&data));
                        data
                    };
                    images.push(EmbeddedImage {
                        picture_name: non_visual.name.clone(),
                        description: non_visual.description.clone(),
                        rel_id: rel_id.clone(),
                        format,
                        data,
                    });
                }
                continue;
            }
            let Object::GraphicFrame(frame) = &anchor.object else {
                continue;
            };
            let Some(rel_id) = &frame.rel_id else {
                continue;
            };
            let relationship = drawing_part.rels().get(rel_id).ok_or_else(|| {
                crate::package::error::Error::Unrecognized {
                    typ: "Drawings part".to_string(),
                    val: format!(
                        "graphic frame {:?} relationship {rel_id:?} is missing",
                        frame.non_visual.name
                    ),
                }
            })?;
            if !matches!(
                relationship.reltype(),
                relationship_type::CHART | relationship_type::STRICT_CHART
            ) {
                continue;
            }
            if relationship.is_external() {
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "Drawings part".to_string(),
                    val: format!("chart relationship {rel_id:?} is external"),
                });
            }
            let chart_part = self.package.get_part(&relationship.target_partname()?)?;
            let graph =
                crate::package::chart_resources::parse_chart_resources(&self.package, chart_part)?;
            charts.push(EmbeddedChart {
                frame_name: frame.non_visual.name.clone(),
                rel_id: rel_id.clone(),
                chart: graph.chart,
                external_data_part: graph.external_data_part,
                user_shapes_part: graph.user_shapes_part,
                additional_relationships: graph.additional_relationships,
            });
        }
        Ok(SheetDrawing {
            sheet_index,
            drawing,
            charts,
            images,
            shapes,
        })
    }

    /// Workbook style table loaded from `xl/styles.bin`.
    pub fn save<W: Write + Seek>(&self, writer: W) -> Result<()> {
        self.package.to_stream(writer)?;
        Ok(())
    }

    pub fn new<R: Read + Seek>(reader: R) -> Result<Self> {
        let package = OpcPackage::from_reader(reader)?;
        Self::from_opc_package(package)
    }

    /// Create an XLSB workbook from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `package` - An already-parsed OPC package
    pub fn from_opc_package(package: OpcPackage) -> Result<Self> {
        let mut workbook = Workbook {
            package,
            worksheets: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            formula_context: Context::default(),
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calc: Props::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
            structured_tables: Vec::new(),
            chart_sheets: Vec::new(),
            sheet_drawings: Vec::new(),
            connections: None,
        };

        workbook.load_workbook_info()?;
        workbook.load_styles()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    fn load_workbook_info(&mut self) -> Result<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;

        let blob = workbook_part.blob();
        let mut iter = Records::new(blob);
        let info = Self::read_workbook(&mut iter)?;
        let external_link_uris = info
            .external_link_rel_ids
            .iter()
            .map(|rel_id| {
                let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
                    crate::package::error::Error::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is missing"
                    ))
                })?;
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is external"
                    )));
                }
                if !matches!(
                    relationship.reltype(),
                    relationship_type::EXTERNAL_LINK | relationship_type::STRICT_EXTERNAL_LINK
                ) {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} has invalid type {:?}",
                        relationship.reltype()
                    )));
                }
                relationship.target_partname().map_err(Into::into)
            })
            .collect::<Result<Vec<_>>>()?;
        let external_books = external_link_uris
            .iter()
            .map(|uri| self.load_external_book(uri))
            .collect::<Result<Vec<_>>>()?;
        let pivot_cache_ids = Self::parse_pivot_cache_ids(workbook_part.blob())?;
        let mut pivot_cache_definitions = Vec::with_capacity(pivot_cache_ids.len());
        for (cache_id, rel_id) in &pivot_cache_ids {
            let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(format!(
                    "PivotCache {cache_id} relationship {rel_id:?} is missing"
                ))
            })?;
            if relationship.is_external()
                || !relationship
                    .reltype()
                    .to_ascii_lowercase()
                    .ends_with("/pivotcachedefinition")
            {
                return Err(crate::package::error::Error::InvalidFormula(format!(
                    "PivotCache {cache_id} relationship is external or has the wrong type"
                )));
            }
            let part = self.package.get_part(&relationship.target_partname()?)?;
            let definition = crate::package::pivot::parse_pivot_cache_definition(part.blob())?;
            pivot_cache_definitions.push((*cache_id, definition));
        }

        let connections = crate::package::connections::package::load_from_workbook(
            &self.package,
            workbook_part.partname(),
        )?;

        let mut tables = Vec::new();
        let mut pivot_views = Vec::new();
        let mut structured_tables = Vec::new();
        let mut chart_sheets = Vec::new();
        let mut sheet_drawings = Vec::new();
        for (sheet_index, rel_id) in info.worksheet_rel_ids.iter().enumerate() {
            let Some(rel_id) = rel_id else { continue };
            let Some(sheet_relationship) = workbook_part.rels().get(rel_id) else {
                continue;
            };
            if sheet_relationship.is_external() {
                continue;
            }
            let sheet_part = self
                .package
                .get_part(&sheet_relationship.target_partname()?)?;
            if CHART_SHEET_RELATIONSHIP_TYPES.contains(&sheet_relationship.reltype()) {
                // Chart Sheet part (MS-XLSB 2.1.7.7): a BIFF12 stream. The
                // chart itself lives in the linked XML Drawings/Chart parts.
                let name = info
                    .worksheet_names
                    .get(sheet_index)
                    .cloned()
                    .ok_or_else(|| crate::package::error::Error::Unrecognized {
                        typ: "BrtBundleSh".to_string(),
                        val: format!("chart sheet index {sheet_index} out of bounds"),
                    })?;
                let state = info.worksheet_states.get(sheet_index).copied().unwrap_or(0);
                let chart_sheet = crate::package::chartsheet::parse_chart_sheet_part(
                    sheet_part.blob(),
                    name,
                    state,
                )?;
                if let Some(drawing_rel_id) = chart_sheet.drawing_rel_id.clone() {
                    let relationship = sheet_part.rels().get(&drawing_rel_id).ok_or_else(|| {
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtDrawing".to_string(),
                            val: format!(
                                "relationship {drawing_rel_id:?} on chart sheet {sheet_index} is missing"
                            ),
                        }
                    })?;
                    if !matches!(
                        relationship.reltype(),
                        relationship_type::DRAWING | relationship_type::STRICT_DRAWING
                    ) || relationship.is_external()
                    {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtDrawing".to_string(),
                            val: format!(
                                "relationship {drawing_rel_id:?} on chart sheet {sheet_index} is external or has the wrong type"
                            ),
                        });
                    }
                    let drawing_part = self.package.get_part(&relationship.target_partname()?)?;
                    sheet_drawings.push(self.load_sheet_drawing(sheet_index, drawing_part)?);
                }
                chart_sheets.push((sheet_index, chart_sheet));
            } else {
                // Worksheet Drawings parts (MS-XLSB 2.1.7.23) are standard
                // SpreadsheetDrawing XML parts discovered through the sheet's
                // drawing relationships.
                for relationship in sheet_part.rels().iter().filter(|relationship| {
                    matches!(
                        relationship.reltype(),
                        relationship_type::DRAWING | relationship_type::STRICT_DRAWING
                    )
                }) {
                    if relationship.is_external() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "worksheet drawing relationship".to_string(),
                            val: "external Drawings part".to_string(),
                        });
                    }
                    let drawing_part = self.package.get_part(&relationship.target_partname()?)?;
                    sheet_drawings.push(self.load_sheet_drawing(sheet_index, drawing_part)?);
                }
            }
            for table_rel_id in crate::package::table::parse_table_part_rel_ids(sheet_part.blob())?
            {
                let relationship = sheet_part.rels().get(&table_rel_id).ok_or_else(|| {
                    crate::package::error::Error::InvalidFormula(format!(
                        "BrtListPart relationship {table_rel_id:?} on sheet {sheet_index} is missing"
                    ))
                })?;
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "BrtListPart relationship {table_rel_id:?} on sheet {sheet_index} is external"
                    )));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let table = crate::package::table::parse_table_part(part.blob())?;
                structured_tables.push((sheet_index, table));
            }
            for relationship in sheet_part.rels().iter().filter(|relationship| {
                matches!(
                    relationship.reltype(),
                    relationship_type::TABLE | relationship_type::STRICT_TABLE
                ) || relationship.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
            }) {
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "worksheet has an external table relationship".to_string(),
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let table = Self::parse_table_definition(part.blob(), sheet_index)?;
                if tables.iter().any(|existing: &TableDefinition| {
                    existing.table_id() == table.table_id()
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "duplicate workbook table ID {}",
                        table.table_id()
                    )));
                }
                if tables.iter().any(|existing: &TableDefinition| {
                    excel_name_eq(existing.display_name(), table.display_name())
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "duplicate workbook table display name {:?}",
                        table.display_name()
                    )));
                }
                tables.push(table);
            }
            for relationship in sheet_part.rels().iter().filter(|relationship| {
                relationship
                    .reltype()
                    .to_ascii_lowercase()
                    .ends_with("/pivottable")
            }) {
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "worksheet has an external PivotTable relationship".to_string(),
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let view = Self::parse_pivot_view(part.blob(), sheet_index)?;
                if !pivot_cache_ids
                    .iter()
                    .any(|(cache_id, _)| *cache_id == view.cache_id())
                {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "PivotTable view {:?} references unknown cache {}",
                        view.name(),
                        view.cache_id()
                    )));
                }
                if pivot_views.iter().any(|existing: &View| {
                    existing.cache_id() == view.cache_id()
                        && existing.sheet_index() == view.sheet_index()
                        && excel_name_eq(existing.name(), view.name())
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "duplicate PivotTable view {:?} for cache {} on sheet {sheet_index}",
                        view.name(),
                        view.cache_id()
                    )));
                }
                pivot_views.push(view);
            }
        }
        self.formula_context = Context {
            worksheet_names: info.worksheet_names.into(),
            supporting_links: info.supporting_links.into(),
            external_sheets: info.external_sheets.into(),
            external_books: external_books.into(),
            defined_names: info.defined_names.into(),
            tables: tables.into(),
            pivot_views: pivot_views.into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        self.worksheet_rel_ids = info.worksheet_rel_ids;
        self.is_1904 = info.is_1904;
        self.calc = info.calc.unwrap_or_default();
        self.pivot_cache_definitions = pivot_cache_definitions;
        self.connections = connections;
        self.structured_tables = structured_tables;
        self.chart_sheets = chart_sheets;
        self.sheet_drawings = sheet_drawings;

        Ok(())
    }

    /// Load shared strings from xl/sharedStrings.bin
    fn load_shared_strings(&mut self) -> Result<()> {
        let shared_strings_uri = litchi_opc::PackURI::new("/xl/sharedStrings.bin")?;
        if let Ok(shared_strings_part) = self.package.get_part(&shared_strings_uri) {
            let blob = shared_strings_part.blob();
            let mut iter = Records::new(blob);
            Self::read_shared_strings(&mut iter, &mut self.shared_strings)?;
        }

        Ok(())
    }

    /// Load workbook styles. The default table keeps style index zero usable
    /// for minimal producer files that omit the optional styles part.
    fn load_styles(&mut self) -> Result<()> {
        let styles_uri = litchi_opc::PackURI::new("/xl/styles.bin")?;
        if let Ok(styles_part) = self.package.get_part(&styles_uri) {
            self.styles = StylesTable::from_bytes(styles_part.blob())?;
        }
        Ok(())
    }
}
