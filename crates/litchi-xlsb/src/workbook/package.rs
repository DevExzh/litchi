//! OPC/package integration for the typed XLSB workbook.

use super::catalog::TabMetadata;
use super::model::{DrawingLoadPolicy, Workbook};
use crate::calc::Props;
use crate::cell_values;
use crate::cell_watches;
use crate::external_link::{Budget, ExternalLinkLimits};
use crate::leniency::{Leniency, Tolerance, ToleranceReport};
use crate::package::error::Result;
use crate::package::formula::{Context, View, excel_name_eq, table::Definition as TableDefinition};
use crate::package::styles_table::StylesTable;
#[cfg(feature = "vba-inspection")]
use crate::package::vba_project::{VbaProject, discover_vba_project};
use crate::package::web_extension_bindings::PackageAppRefs;
use crate::raw::Records;
use crate::sheet::{Kind, State};
use crate::sparkline;
use litchi_ooxml_common::embedded;
use litchi_ooxml_common::ribbon;
use litchi_ooxml_common::web;
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::part::Part;
use litchi_opc::{OpcError, OpcPackage};
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek, Write};
use std::sync::Arc;

/// Chart sheet relationship types documented by MS-XLSB 2.1.7.7.
const CHART_SHEET_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
];

const DIALOG_SHEET_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
];

const MACRO_SHEET_RELATIONSHIP_TYPES: &[&str] = &[
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
];

const XLSB_WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";
const XLSB_CHARTSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.chartsheet";
const XLSB_DIALOGSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.dialogsheet";
const XLSB_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.macrosheet";
const XLSB_INTL_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.intlmacrosheet";
const XLSB_SHARED_STRINGS_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";

fn is_worksheet_relationship(reltype: &str) -> bool {
    matches!(
        reltype,
        relationship_type::WORKSHEET | relationship_type::STRICT_WORKSHEET
    )
}

fn is_known_non_worksheet_relationship(reltype: &str) -> bool {
    CHART_SHEET_RELATIONSHIP_TYPES.contains(&reltype)
        || DIALOG_SHEET_RELATIONSHIP_TYPES.contains(&reltype)
        || MACRO_SHEET_RELATIONSHIP_TYPES.contains(&reltype)
}

fn validate_sheet_content_type(part: &dyn Part, kind: Kind) -> Result<()> {
    let expected = match kind {
        Kind::Worksheet => XLSB_WORKSHEET_CONTENT_TYPE,
        Kind::Chart => XLSB_CHARTSHEET_CONTENT_TYPE,
        Kind::Dialog => XLSB_DIALOGSHEET_CONTENT_TYPE,
        Kind::Macro => XLSB_MACROSHEET_CONTENT_TYPE,
        Kind::IntlMacro => XLSB_INTL_MACROSHEET_CONTENT_TYPE,
        Kind::Unbound => {
            return Err(crate::package::error::Error::UnsupportedFeature(
                "unbound XLSB tabs have no physical content type",
            ));
        },
    };
    if part.content_type() != expected {
        return Err(crate::package::error::Error::InvalidContentType {
            expected: "a known value",
        });
    }
    Ok(())
}

impl Workbook {
    /// Read the optional workbook Theme part under this projection's
    /// [`crate::SemanticLimits::theme_value`].
    pub fn theme(&self) -> Result<Option<crate::theme::Snapshot>> {
        crate::theme::read(&self.package, self.semantic_limits.theme_value())
    }

    /// Read and bind the Worksheet Binary Index attached to one worksheet.
    /// Offsets are validated against the worksheet bytes during this call;
    /// the returned metadata owns the parsed index bytes but does not retain a
    /// live view of the workbook's worksheet part.
    pub fn worksheet_binary_index(
        &self,
        worksheet_index: usize,
    ) -> Result<Option<crate::binary_index::WorksheetBinaryIndex>> {
        let limits = self.semantic_limits.binary_index_value();
        let worksheet_uri = self.worksheet_uri(worksheet_index)?;
        let worksheet = self.package.get_part(&worksheet_uri)?;
        let mut relationships = worksheet.rels().iter().filter(|relationship| {
            relationship.reltype() == crate::binary_index::BINARY_INDEX_RELATIONSHIP
        });
        let Some(relationship) = relationships.next() else {
            return Ok(None);
        };
        if relationships.next().is_some() {
            return Err(crate::package::error::Error::InvalidRelationship(
                "worksheet has multiple binary index relationships",
            ));
        }
        if relationship.is_external() {
            return Err(crate::package::error::Error::InvalidRelationship(
                "worksheet binary index relationship is external",
            ));
        }
        let index_uri = relationship.target_partname()?;
        let index_part = self.package.get_part(&index_uri)?;
        if index_part.content_type() != crate::binary_index::BINARY_INDEX_CONTENT_TYPE {
            return Err(crate::package::error::Error::InvalidContentType {
                expected: crate::binary_index::BINARY_INDEX_CONTENT_TYPE,
            });
        }
        if !index_part.rels().is_empty() {
            return Err(crate::package::error::Error::InvalidRelationship(
                "worksheet binary index part must not have relationships",
            ));
        }
        crate::binary_index::WorksheetBinaryIndex::from_parts(
            index_part.blob(),
            worksheet.blob(),
            limits,
        )
        .map(Some)
    }

    /// Reparse an owned candidate while retaining this workbook's typed
    /// drawing projection boundary.
    pub(crate) fn reparse_candidate(&self, package: OpcPackage) -> Result<Self> {
        Self::reparse_candidate_with_policy(
            package,
            self.external_link_limits,
            self.drawing_load_policy,
            self.leniency,
            self.raw_limits,
        )
    }

    /// Reparse bytes after a caller has selected the projection policies.
    /// This is the single internal seam used by detached CRUD validators so
    /// opaque drawing parts are never parsed accidentally.
    ///
    /// Raw OPC callers pass [`Leniency::Strict`] to retain the historical
    /// strict worksheet projection; workbook-owned callers pass their
    /// retained policy before validating lazy worksheet projections.
    pub(crate) fn reparse_candidate_with_policy(
        package: OpcPackage,
        external_link_limits: ExternalLinkLimits,
        drawing_load_policy: DrawingLoadPolicy,
        leniency: Leniency,
        raw_limits: crate::raw::Limits,
    ) -> Result<Self> {
        let mut workbook = Self::from_opc_package_with_external_link_limits_and_drawing_policy(
            package,
            external_link_limits,
            drawing_load_policy,
            raw_limits,
        )?;
        workbook.set_projection_leniency(leniency);
        Ok(workbook)
    }

    /// Read source-bound, editable stored values on one worksheet under this
    /// projection's [`crate::SemanticLimits::cell_values_value`].
    ///
    /// This inventories ordinary scalar cells, formula cached results, and
    /// inert formula payloads for structural editing.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid worksheet selector or malformed stream.
    pub fn cell_values(&self, worksheet_index: usize) -> Result<cell_values::Snapshot> {
        let uri = self.worksheet_uri(worksheet_index)?;
        cell_values::workbook::read_with_limits_and_leniency(
            &self.package,
            &uri,
            self.semantic_limits.cell_values_value(),
            self.leniency,
        )
    }

    /// Read editable cell values from a worksheet selected by exact name.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid worksheet name or malformed stream.
    pub fn cell_values_by_name(&self, worksheet_name: &str) -> Result<cell_values::Snapshot> {
        self.cell_values(self.worksheet_index(worksheet_name)?)
    }

    /// Read the optional sparkline groups of the worksheet at worksheet
    /// ordinal `worksheet_index` under this projection's
    /// [`crate::SemanticLimits::sparklines_value`].
    pub fn sparklines(&self, worksheet_index: usize) -> Result<sparkline::Snapshot> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let snapshot = sparkline::workbook::read_with_limits(
            &self.package,
            &uri,
            self.semantic_limits.sparklines_value(),
        )?;
        sparkline::workbook::validate_context(&snapshot, &self.formula_context)?;
        Ok(snapshot)
    }

    /// Read optional sparkline groups selected by worksheet name.
    pub fn sparklines_by_name(&self, worksheet_name: &str) -> Result<sparkline::Snapshot> {
        let index = self.worksheet_index(worksheet_name)?;
        self.sparklines(index)
    }

    /// Read the typed cell-watch and worksheet phonetic snapshot selected by
    /// zero-based worksheet index.
    pub fn cell_watches(&self, worksheet_index: usize) -> Result<cell_watches::Snapshot> {
        let uri = self.worksheet_uri(worksheet_index)?;
        cell_watches::workbook::read(&self.package, &uri, self.raw_limits)
    }

    /// Read the typed cell-watch and phonetic snapshot selected by worksheet
    /// name.
    pub fn cell_watches_by_name(&self, worksheet_name: &str) -> Result<cell_watches::Snapshot> {
        let index = self.worksheet_index(worksheet_name)?;
        self.cell_watches(index)
    }

    /// Load inert persisted Office Add-in task panes.
    pub fn task_panes(&self) -> Result<Option<web::Panes>> {
        Ok(web::load(&self.package, &web::Limits::standard())?)
    }

    pub(crate) fn validate_task_pane_bindings(&self, panes: &web::Panes) -> Result<()> {
        let package_refs = PackageAppRefs::new(
            panes
                .iter()
                .flat_map(|pane| pane.add_in().bindings().iter()),
        )?;
        for index in 0..self.worksheet_names.len() {
            let worksheet = self.cached_worksheet(index)?;
            package_refs.validate(worksheet.web_extension_bindings())?;
        }
        Ok(())
    }

    /// Read the bounded, inert package-level Ribbon customizations.
    pub fn ribbon(&self) -> Result<ribbon::Set<'_>> {
        Ok(ribbon::load(&self.package, &ribbon::Limits::standard())?)
    }

    /// Discover inert embedded-object and embedded-package relationships
    /// using the shared safe default resource limits.
    ///
    /// Use [`embedded::scan`] with [`Self::opc_package`] when a lower
    /// layer needs explicitly tuned limits.
    pub fn embedded(&self) -> Result<Vec<embedded::Entry<'_>>> {
        Ok(embedded::scan(&self.package, &embedded::Limits::default())?)
    }

    /// Get the underlying OPC package.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.package
    }

    /// Consume this workbook and return the OPC package it publishes.
    ///
    /// This is the seam that lets a validated candidate parse be unwrapped
    /// back into a package for callers that own a package rather than a
    /// workbook.
    pub(crate) fn into_opc_package(self) -> OpcPackage {
        self.package
    }

    /// Get mutable OPC access for XLSB-internal package adapters.
    ///
    /// Public callers must use [`Self::edit_opc`], which stages a structural
    /// candidate and reparses the complete XLSB host before publication. The
    /// signature state is left as it is: a changed signed package is refused
    /// when it is published.
    #[allow(
        dead_code,
        reason = "retained for BIFF12 codec completeness and staged host integration"
    )]
    pub(crate) fn opc_package_mut(&mut self) -> &mut OpcPackage {
        &mut self.package
    }

    /// Run one graph change on a candidate copy of this workbook's OPC
    /// graph, validate the candidate and adopt it (the crate-internal seam
    /// behind the remaining in-place workbook paths). An error leaves this
    /// workbook unchanged.
    pub(crate) fn edit_opc<T>(
        &mut self,
        edit: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let mut candidate = self.package.clone();
        let value = edit(&mut candidate)?;
        candidate.validate_signature_edit_from(&self.package)?;

        Self::validate_edit_candidate(&candidate)?;
        let validated = self.reparse_candidate(candidate)?;
        *self = validated;
        Ok(value)
    }

    pub(crate) fn validate_edit_candidate(package: &OpcPackage) -> Result<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook = package.get_part(&workbook_uri)?;
        if workbook.content_type() != content_type::XLSB_BIN {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "XLSB workbook content type",
                val: "expected, found",
            });
        }

        let main = package.main_document_part()?;
        if main.partname() != workbook.partname() || main.content_type() != content_type::XLSB_BIN {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "XLSB main workbook relationship",
                val: "expected to be the binary workbook main part",
            });
        }

        // These readers validate relationship cardinality, target mode,
        // target content type, and orphan/inbound graph invariants without
        // parsing or executing opaque payload bytes.
        #[cfg(feature = "vba-inspection")]
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
    #[cfg(feature = "sign")]
    pub fn signatures(&self) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.package.signatures()
    }

    /// Verify package signatures with an explicit trust-neutral policy.
    #[cfg(feature = "sign")]
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> litchi_opc::sign::Result<Vec<litchi_opc::sign::Report>> {
        self.package.signatures_with(policy)
    }

    /// Discover the attached MS-XLSB VBA project and declared signatures.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, verify, or execute VBA project or
    /// signature bytes.
    #[cfg(feature = "vba-inspection")]
    pub fn vba(&self) -> Result<Option<VbaProject>> {
        let workbook = self.package.main_document_part()?;
        discover_vba_project(&self.package, workbook)
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

    fn load_sheet_drawing(
        &self,
        sheet_index: usize,
        drawing_part: &dyn Part,
    ) -> Result<crate::package::drawing::SheetDrawing> {
        use crate::package::drawing::{EmbeddedChart, EmbeddedImage, Object, SheetDrawing};
        let drawing_xml = std::str::from_utf8(drawing_part.blob()).map_err(|_error| {
            crate::package::error::Error::Encoding("Drawings part is not UTF-8")
        })?;
        let shape_limits = crate::shapes::Limits::from_core(self.semantic_limits.xml_parts_value());
        let shapes = crate::shapes::read(drawing_xml, &shape_limits)?.unwrap_or_default();
        let drawing = crate::package::drawing::parse_drawing_part(
            drawing_part.blob(),
            self.semantic_limits.xml_parts_value(),
        )?;
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
                let relationship = drawing_part.rels().get(rel_id).ok_or({
                    crate::package::error::Error::Unrecognized {
                        typ: "Drawings part",
                        val: "picture relationship is missing",
                    }
                })?;
                if matches!(
                    relationship.reltype(),
                    relationship_type::IMAGE | relationship_type::STRICT_IMAGE
                ) {
                    if relationship.is_external() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "Drawings part",
                            val: "image relationship is external",
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
            let relationship = drawing_part.rels().get(rel_id).ok_or({
                crate::package::error::Error::Unrecognized {
                    typ: "Drawings part",
                    val: "graphic frame relationship is missing",
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
                    typ: "Drawings part",
                    val: "chart relationship is external",
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
        self.package.write_to(writer)?;
        Ok(())
    }

    /// Read and validate an XLSB workbook with the default bounded OPC limits.
    pub fn new<R: Read + Seek>(reader: R) -> Result<Self> {
        Self::new_with_options(reader, crate::OpenOptions::default())
    }

    /// Read and validate an XLSB workbook under explicit ingress policies.
    ///
    /// With [`crate::OpenOptions::parse_drawings`] `false` the projection
    /// retains every drawing, chart, image and relationship part for lossless
    /// save but does not populate the typed [`Workbook::sheet_drawings`]
    /// inventory.
    ///
    /// # Errors
    ///
    /// Returns [`crate::package::PackageError::UnsupportedOpenPolicy`] for a
    /// cache or execution policy (positional-source only), the OPC limit and
    /// archive errors, and the workbook validation errors.
    pub fn new_with_options<R: Read + Seek>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let package = crate::package::ingress::from_reader(reader, options.source_limits)?;
        Self::from_opc_package_with_options(package, options)
    }

    /// Create an XLSB workbook from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    pub fn from_opc_package(package: OpcPackage) -> Result<Self> {
        Self::from_opc_package_with_options(package, crate::OpenOptions::default())
    }

    /// Create an XLSB workbook from an already-parsed OPC package under
    /// explicit semantic policies: external-link limits, leniency and the
    /// drawing policy. The package's own source limits already applied when
    /// it was read.
    ///
    /// The Excel-compatible repairs are applied under every leniency; see
    /// [`crate::leniency`]. [`Leniency::TolerateCachedNumbers`] additionally
    /// repairs non-normal cached numbers in materialized worksheets and
    /// cell-value snapshots, recording each repair in the report of the
    /// object that read the cell.
    ///
    /// # Errors
    ///
    /// Returns [`crate::package::PackageError::UnsupportedOpenPolicy`] for a
    /// cache or execution policy, and the workbook validation errors.
    pub fn from_opc_package_with_options(
        package: OpcPackage,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_owning()?;
        let mut workbook = Self::from_opc_package_with_external_link_limits_and_drawing_policy(
            package,
            options.external_link_limits,
            options.drawing_load_policy(),
            options.raw_limits,
        )?;
        workbook.leniency = options.leniency;
        workbook.semantic_limits = options.semantic_limits;
        Ok(workbook)
    }

    pub(crate) fn from_opc_package_with_external_link_limits(
        package: OpcPackage,
        external_link_limits: ExternalLinkLimits,
        raw_limits: crate::raw::Limits,
    ) -> Result<Self> {
        Self::from_opc_package_with_external_link_limits_and_drawing_policy(
            package,
            external_link_limits,
            DrawingLoadPolicy::Eager,
            raw_limits,
        )
    }

    pub(crate) fn from_opc_package_with_external_link_limits_and_drawing_policy(
        package: OpcPackage,
        external_link_limits: ExternalLinkLimits,
        drawing_load_policy: DrawingLoadPolicy,
        raw_limits: crate::raw::Limits,
    ) -> Result<Self> {
        let mut external_link_budget = external_link_limits.budget();
        // The packaging repairs of the OPC read come first in discovery order.
        let tolerance = ToleranceReport::from_package(package.tolerance_report());
        let mut workbook = Workbook {
            package,
            worksheet_cache: super::cache::WorksheetCache::new(),
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            worksheet_positions: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            sheet_tabs: Vec::new(),
            active_catalog_position: None,
            formula_context: Context::default(),
            raw_limits,
            external_link_limits,
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calc: Props::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
            structured_tables: Vec::new(),
            chart_sheets: Vec::new(),
            sheet_drawings: Vec::new(),
            drawing_load_policy,
            connections: None,
            tolerance,
            leniency: Leniency::Strict,
            semantic_limits: crate::SemanticLimits::new(),
        };

        workbook.load_workbook_info(
            &mut external_link_budget,
            drawing_load_policy == DrawingLoadPolicy::Eager,
        )?;
        workbook.load_styles()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    fn load_workbook_info(
        &mut self,
        external_link_budget: &mut Budget,
        load_drawings: bool,
    ) -> Result<()> {
        let workbook_part = self.package.main_document_part()?;

        let blob = workbook_part.blob();
        let mut iter = Records::new(blob, self.raw_limits);
        let info = Self::read_workbook(&mut iter, &mut self.tolerance, self.raw_limits)?;
        if info.worksheet_names.len() != info.worksheet_rel_ids.len() {
            return Err(crate::package::error::Error::InvalidFormat(
                "XLSB sheet names and relationship identifiers have different lengths",
            ));
        }
        if info.worksheet_names.len() != info.worksheet_states.len() {
            return Err(crate::package::error::Error::InvalidFormat(
                "XLSB sheet names and visibility states have different lengths",
            ));
        }
        let active_catalog_position =
            (!info.worksheet_names.is_empty()).then_some(info.active_catalog_position.unwrap_or(0));
        // BrtSupBookSrc identifiers are workbook-stream metadata. The
        // external-link budget starts when each referenced part is parsed.
        external_link_budget.preflight_links(info.external_link_rel_ids.len())?;
        let mut external_link_uris = Vec::new();
        external_link_uris
            .try_reserve_exact(info.external_link_rel_ids.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB external-link URI map",
                source,
            })?;
        for rel_id in &info.external_link_rel_ids {
            let relationship = workbook_part.rels().get(rel_id).ok_or({
                crate::package::error::Error::InvalidFormula(
                    "BrtSupBookSrc relationship is missing",
                )
            })?;
            if relationship.is_external() {
                return Err(crate::package::error::Error::InvalidFormula(
                    "BrtSupBookSrc relationship is external",
                ));
            }
            if !matches!(
                relationship.reltype(),
                relationship_type::EXTERNAL_LINK | relationship_type::STRICT_EXTERNAL_LINK
            ) {
                return Err(crate::package::error::Error::InvalidFormula(
                    "BrtSupBookSrc relationship has invalid type",
                ));
            }
            external_link_uris.push(relationship.target_partname()?);
        }
        let mut external_books = Vec::new();
        external_books
            .try_reserve_exact(external_link_uris.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB external-link result map",
                source,
            })?;
        for uri in &external_link_uris {
            external_books.push(self.load_external_book_with_budget(uri, external_link_budget)?);
        }
        let pivot_cache_ids = Self::parse_pivot_cache_ids(workbook_part.blob(), self.raw_limits)?;
        let mut pivot_cache_definitions = Vec::with_capacity(pivot_cache_ids.len());
        for (cache_id, rel_id) in &pivot_cache_ids {
            let relationship = workbook_part.rels().get(rel_id).ok_or(
                crate::package::error::Error::InvalidFormula("PivotCache relationship is missing"),
            )?;
            if relationship.is_external()
                || !relationship
                    .reltype()
                    .to_ascii_lowercase()
                    .ends_with("/pivotcachedefinition")
            {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotCache relationship is external or has the wrong type",
                ));
            }
            let part = self.package.get_part(&relationship.target_partname()?)?;
            let definition = crate::package::pivot::parse_pivot_cache_definition_recording(
                part.blob(),
                &mut self.tolerance,
                self.raw_limits,
            )?;
            pivot_cache_definitions.push((*cache_id, definition));
        }

        let connections = crate::package::connections::package::load_from_workbook_recording(
            &self.package,
            workbook_part.partname(),
            &mut self.tolerance,
            self.raw_limits,
        )?;

        let mut tables = Vec::new();
        let mut pivot_views = Vec::new();
        let mut structured_tables = Vec::new();
        let mut chart_sheets = Vec::new();
        let mut sheet_drawings = Vec::new();
        let mut worksheet_positions = Vec::new();
        worksheet_positions
            .try_reserve_exact(info.worksheet_rel_ids.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB worksheet position map",
                source,
            })?;
        let mut worksheet_names = Vec::new();
        worksheet_names
            .try_reserve_exact(info.worksheet_rel_ids.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB worksheet-name map",
                source,
            })?;
        let mut sheet_tabs = Vec::new();
        sheet_tabs
            .try_reserve_exact(info.worksheet_rel_ids.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB sheet-tab catalog",
                source,
            })?;
        let mut sheet_targets = HashSet::new();
        sheet_targets
            .try_reserve(info.worksheet_rel_ids.len())
            .map_err(|source| crate::package::error::Error::Allocation {
                resource: "eager XLSB sheet-target validation",
                source,
            })?;
        for (catalog_position, rel_id) in info.worksheet_rel_ids.iter().enumerate() {
            let name = &info.worksheet_names[catalog_position];
            let state = State::from_hs_state(info.worksheet_states[catalog_position])?;
            let Some(rel_id) = rel_id.as_deref() else {
                if state != State::VeryHidden {
                    return Err(crate::package::error::Error::InvalidRelationship(
                        "unbound XLSB sheet tabs must be very hidden",
                    ));
                }
                sheet_tabs.push(TabMetadata::new(name.to_owned(), Kind::Unbound, state));
                continue;
            };
            let relationship = workbook_part.rels().get(rel_id).ok_or({
                crate::package::error::Error::Unrecognized {
                    typ: "sheet relationship",
                    val: "relationship for sheet catalog position is missing",
                }
            })?;
            if relationship.is_external() {
                return Err(crate::package::error::Error::UnsupportedFeature(
                    "sheet catalog position has an external relationship",
                ));
            }
            let Some(kind) = Kind::from_relationship_type(relationship.reltype()) else {
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "sheet relationship",
                    val: "relationship for sheet catalog position has unsupported type",
                });
            };
            let partname = relationship.target_partname()?;
            if !sheet_targets.insert(partname.as_str().to_ascii_lowercase()) {
                return Err(crate::package::error::Error::InvalidRelationship(
                    "multiple XLSB sheet tabs resolve to one physical part",
                ));
            }
            let sheet_part = self.package.get_part(&partname)?;
            validate_sheet_content_type(sheet_part, kind)?;
            sheet_tabs.push(TabMetadata::new(name.to_owned(), kind, state));
            if kind.is_worksheet() {
                worksheet_positions.push(catalog_position);
                worksheet_names.push(name.to_owned());
            }
        }
        for (sheet_index, rel_id) in info.worksheet_rel_ids.iter().enumerate() {
            let Some(rel_id) = rel_id.as_deref() else {
                continue;
            };
            let sheet_relationship = workbook_part.rels().get(rel_id).ok_or({
                crate::package::error::Error::Unrecognized {
                    typ: "sheet relationship",
                    val: "relationship for sheet catalog position is missing",
                }
            })?;
            if sheet_relationship.is_external() {
                return Err(crate::package::error::Error::UnsupportedFeature(
                    "sheet catalog position has an external relationship",
                ));
            }
            let reltype = sheet_relationship.reltype();
            if !is_worksheet_relationship(reltype)
                && !CHART_SHEET_RELATIONSHIP_TYPES.contains(&reltype)
            {
                if is_known_non_worksheet_relationship(reltype) {
                    continue;
                }
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "sheet relationship",
                    val: "relationship for sheet catalog position has unsupported type",
                });
            }
            let sheet_part = self
                .package
                .get_part(&sheet_relationship.target_partname()?)?;
            if CHART_SHEET_RELATIONSHIP_TYPES.contains(&sheet_relationship.reltype()) {
                // Chart Sheet part (MS-XLSB 2.1.7.7): a BIFF12 stream. The
                // chart itself lives in the linked XML Drawings/Chart parts.
                let name = info.worksheet_names.get(sheet_index).cloned().ok_or(
                    crate::package::error::Error::Unrecognized {
                        typ: "BrtBundleSh",
                        val: "chart sheet index out of bounds",
                    },
                )?;
                let state = info.worksheet_states[sheet_index];
                let chart_sheet = crate::package::chartsheet::parse_chart_sheet_part(
                    sheet_part.blob(),
                    name,
                    state,
                    self.raw_limits,
                )?;
                if let Some(drawing_rel_id) = chart_sheet.drawing_rel_id.clone() {
                    let relationship = sheet_part.rels().get(&drawing_rel_id).ok_or({
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtDrawing",
                            val: "relationship on chart sheet is missing",
                        }
                    })?;
                    if !matches!(
                        relationship.reltype(),
                        relationship_type::DRAWING | relationship_type::STRICT_DRAWING
                    ) || relationship.is_external()
                    {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtDrawing",
                            val: "relationship on chart sheet is external or has the wrong type",
                        });
                    }
                    let drawing_part = self.package.get_part(&relationship.target_partname()?)?;
                    if load_drawings {
                        sheet_drawings.push(self.load_sheet_drawing(sheet_index, drawing_part)?);
                    }
                }
                chart_sheets.push((sheet_index, chart_sheet));
                continue;
            }
            let worksheet_index = worksheet_positions
                .binary_search(&sheet_index)
                .map_err(|_missing_worksheet_position| crate::package::error::Error::Unrecognized {
                    typ: "worksheet catalog",
                    val: "worksheet relationship at catalog position has no public worksheet ordinal",
                })?;
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
                        typ: "worksheet drawing relationship",
                        val: "external Drawings part",
                    });
                }
                let drawing_part = self.package.get_part(&relationship.target_partname()?)?;
                if load_drawings {
                    sheet_drawings.push(self.load_sheet_drawing(sheet_index, drawing_part)?);
                }
            }
            for table_rel_id in crate::package::table::parse_table_part_rel_ids_recording(
                sheet_part.blob(),
                &mut self.tolerance,
                self.raw_limits,
            )? {
                let relationship = sheet_part.rels().get(&table_rel_id).ok_or({
                    crate::package::error::Error::InvalidFormula(
                        "BrtListPart relationship on sheet is missing",
                    )
                })?;
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "BrtListPart relationship on sheet is external",
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let table = crate::package::table::parse_table_part_recording(
                    part.blob(),
                    &mut self.tolerance,
                    self.raw_limits,
                )?;
                structured_tables.push((worksheet_index, table));
            }
            for relationship in sheet_part.rels().iter().filter(|relationship| {
                matches!(
                    relationship.reltype(),
                    relationship_type::TABLE | relationship_type::STRICT_TABLE
                )
            }) {
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "worksheet has an external table relationship",
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                // The same BrtBeginList/BrtBeginListCol records were already
                // parsed (and any repair recorded) by the typed table parser
                // above; this formula-context pass must not record them twice.
                let mut already_recorded = ToleranceReport::default();
                let table = Self::parse_table_definition(
                    part.blob(),
                    sheet_index,
                    &mut Tolerance::recording(&mut already_recorded),
                    self.raw_limits,
                )?;
                if tables
                    .iter()
                    .any(|existing: &TableDefinition| existing.table_id() == table.table_id())
                {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "duplicate workbook table ID",
                    ));
                }
                if tables.iter().any(|existing: &TableDefinition| {
                    excel_name_eq(existing.display_name(), table.display_name())
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "duplicate workbook table display name",
                    ));
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
                        "worksheet has an external PivotTable relationship",
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let view = Self::parse_pivot_view(part.blob(), sheet_index, self.raw_limits)?;
                if !pivot_cache_ids
                    .iter()
                    .any(|(cache_id, _)| *cache_id == view.cache_id())
                {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "PivotTable view references unknown cache",
                    ));
                }
                if pivot_views.iter().any(|existing: &View| {
                    existing.cache_id() == view.cache_id()
                        && existing.sheet_index() == view.sheet_index()
                        && excel_name_eq(existing.name(), view.name())
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "duplicate PivotTable view for cache on sheet",
                    ));
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
        self.worksheet_names = worksheet_names;
        self.worksheet_positions = worksheet_positions;
        self.worksheet_rel_ids = info.worksheet_rel_ids;
        self.sheet_tabs = sheet_tabs;
        self.active_catalog_position = active_catalog_position;
        self.is_1904 = info.is_1904;
        self.calc = info.calc.unwrap_or_default();
        self.pivot_cache_definitions = pivot_cache_definitions;
        self.connections = connections;
        self.structured_tables = structured_tables;
        self.chart_sheets = chart_sheets;
        self.sheet_drawings = sheet_drawings;

        Ok(())
    }

    /// Load shared strings through the workbook's optional relationship.
    fn load_shared_strings(&mut self) -> Result<()> {
        let shared_strings_uri = {
            let workbook_part = self.package.main_document_part()?;
            let mut found = None;
            for relationship in workbook_part.rels().iter() {
                if !matches!(
                    relationship.reltype(),
                    relationship_type::SHARED_STRINGS | relationship_type::STRICT_SHARED_STRINGS
                ) {
                    continue;
                }
                if found.is_some() {
                    return Err(crate::package::error::Error::InvalidRelationship(
                        "XLSB workbook has multiple shared-string relationships",
                    ));
                }
                if relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidRelationship(
                        "XLSB workbook shared-string relationship is external",
                    ));
                }
                found = Some(relationship.target_partname()?);
            }
            found
        };
        let Some(shared_strings_uri) = shared_strings_uri else {
            return Ok(());
        };
        let shared_strings_part = self.package.get_part(&shared_strings_uri)?;
        if shared_strings_part.content_type() != XLSB_SHARED_STRINGS_CONTENT_TYPE {
            return Err(crate::package::error::Error::InvalidContentType {
                expected: XLSB_SHARED_STRINGS_CONTENT_TYPE,
            });
        }
        let blob = shared_strings_part.blob();
        let mut iter = Records::new(blob, self.raw_limits);
        Self::read_shared_strings(
            &mut iter,
            &mut self.shared_strings,
            &mut self.tolerance,
            self.raw_limits,
        )?;

        Ok(())
    }

    /// Load workbook styles. The default table keeps style index zero usable
    /// for minimal producer files that omit the optional styles part.
    fn load_styles(&mut self) -> Result<()> {
        let styles_uri = litchi_opc::PackURI::new("/xl/styles.bin")?;
        match self.package.get_part(&styles_uri) {
            Ok(styles_part) => {
                self.styles = StylesTable::from_bytes_with_report(
                    styles_part.blob(),
                    self.raw_limits,
                    &mut self.tolerance,
                )?;
            },
            Err(OpcError::PartNotFound(_)) => {},
            Err(error) => return Err(error.into()),
        }
        Ok(())
    }
}
