//! Excel Workbook implementation.
//!
//! This module provides the concrete implementation of the Workbook trait
//! for Excel (.xlsx) files using the Office Open XML format.

use crate::common::DocumentProperties;
use crate::pivot::PivotTable;
use crate::ribbonx::{
    RibbonCustomization, RibbonCustomizationVersion, load_ribbon_customization,
    load_ribbon_customizations, store_ribbon_customization,
};
use crate::web_extensions::{
    OoxmlConformance, WebExtensionTaskPanes, load_web_extension_task_panes,
    remove_web_extension_task_panes, store_web_extension_task_panes,
};
use crate::xlsx::active_x::{
    ActiveXControlSet, load_from_worksheet as load_worksheet_active_x,
    remove_from_worksheet as remove_worksheet_active_x,
    replace_on_worksheet as replace_worksheet_active_x,
    store_on_worksheet as store_worksheet_active_x,
};
use crate::xlsx::calculation_chain::{
    CalculationChain, CalculationChainConformance, load_calculation_chain,
    remove_calculation_chain, store_calculation_chain,
};
use crate::xlsx::calculation_properties::{
    WorkbookCalculationProperties, parse_workbook_calculation_properties,
};
use crate::xlsx::data_validation::{
    DataValidationCollection, parse_data_validation_collections,
    replace_data_validation_collections, validate_data_validation_collections,
};
use crate::xlsx::external_links::{
    ExternalLinkConformance, ExternalLinkEntry, ExternalLinkKind,
    build_external_link_part_with_conformance, load_external_link,
};
use crate::xlsx::named_sheet_view::{
    NamedSheetViews, load_worksheet_named_sheet_views, remove_worksheet_named_sheet_views,
    store_worksheet_named_sheet_views,
};
use crate::xlsx::sheet_protection::{
    WorksheetProtectedRangeCollection, WorksheetProtection, WorksheetProtectionMetadata,
    parse_worksheet_protection, replace_worksheet_protection,
    validate_worksheet_protection_metadata,
};
use crate::xlsx::vba_project::{
    VbaProject, discover_vba_project, remove_vba_project as remove_workbook_vba_project,
    store_vba_project as store_workbook_vba_project,
};
use crate::xlsx::volatile_dependencies::{
    VolatileDependencies, VolatileDependenciesConformance,
    load_from_package_with_conformance as load_volatile_dependencies,
    remove_from_package as remove_volatile_dependencies,
    store_in_package as store_volatile_dependencies,
};
use crate::xlsx::web_extension_bindings::{
    WorksheetWebExtensionBinding, parse_worksheet_web_extension_bindings,
    replace_worksheet_web_extension_bindings as patch_worksheet_web_extension_bindings,
    validate_worksheet_web_extension_apprefs,
};
use crate::xlsx::workbook_protection::{WorkbookProtectionMetadata, parse_workbook_protection};
use crate::xlsx::writer::workbook::{
    generate_pivot_cache_definition_xml, generate_pivot_cache_records_xml,
    generate_pivot_table_definition_xml, render_pivot_table_sheet_cells,
};
use crate::xlsx::writer::{MutableWorkbookData, MutableWorksheet, NamedRange};
use crate::xlsx::xml_maps::{
    XmlMapConformance, XmlMapInfo, load_from_package_with_conformance as load_xml_maps,
    remove_from_package as remove_xml_maps, store_in_package as store_xml_maps,
};
use crate::xlsx::{Cell, SharedStrings, Styles};
use litchi_core::sheet::{
    Result as SheetResult, WorkbookTrait, Worksheet as WorksheetTrait, WorksheetIterator,
};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::{HashMap, HashSet};

use super::parsers::workbook_parser;
use super::worksheet::{
    ArrayFormula, Worksheet, WorksheetInfo, WorksheetIterator as XlsxWorksheetIterator,
};

fn next_active_x_relationship_id(
    occupied: &mut HashSet<String>,
    control_index: usize,
    preview: bool,
) -> String {
    let kind = if preview { "Preview" } else { "Control" };
    let mut suffix = 0usize;
    loop {
        let id = format!("_litchiActiveX{kind}{control_index}_{suffix}");
        if occupied.insert(id.clone()) {
            return id;
        }
        suffix += 1;
    }
}

/// Rekey an index-keyed per-worksheet mutation map after a worksheet
/// removal: the entry for the removed index is dropped and later entries
/// shift one index down.
fn shift_index_keyed_mutations<V>(map: &mut HashMap<usize, V>, removed: usize) {
    map.remove(&removed);
    let shifted = map
        .drain()
        .map(|(index, value)| (if index > removed { index - 1 } else { index }, value))
        .collect::<Vec<_>>();
    map.extend(shifted);
}

/// Extract a numeric stem from a part name (`prefix{digits}suffix`).
fn numeric_part_stem(name: &str, prefix: &str, suffix: &str) -> Option<u32> {
    name.strip_prefix(prefix)?
        .strip_suffix(suffix)?
        .parse::<u32>()
        .ok()
}

/// Whether a package part name belongs to a writer-owned sheet part
/// family but has no live backing sheet in the writer data. Names that
/// do not parse cleanly are conservatively kept.
fn is_stale_sheet_part(
    name: &str,
    worksheet_ids: &HashSet<u32>,
    all_sheet_ids: &HashSet<u32>,
    chartsheet_count: usize,
) -> bool {
    // /xl/worksheets/sheet{id}.xml — regenerated once per worksheet.
    if let Some(id) = numeric_part_stem(name, "/xl/worksheets/sheet", ".xml") {
        return !worksheet_ids.contains(&id);
    }
    // /xl/chartsheets/sheet{n}.xml and its drawing — regenerated per slot.
    if let Some(n) = numeric_part_stem(name, "/xl/chartsheets/sheet", ".xml") {
        return n as usize > chartsheet_count;
    }
    if let Some(n) = numeric_part_stem(name, "/xl/drawings/drawingChartsheet", ".xml") {
        return n as usize > chartsheet_count;
    }
    // /xl/charts/chart{id}_{m}.xml — hosted by a worksheet or chartsheet.
    if let Some(rest) = name
        .strip_prefix("/xl/charts/chart")
        .and_then(|rest| rest.strip_suffix(".xml"))
        && let Some((id, _)) = rest.split_once('_')
        && let Ok(id) = id.parse::<u32>()
    {
        return !all_sheet_ids.contains(&id);
    }
    // /xl/drawings/drawing{id}.xml and vmlDrawing{id}.vml — per worksheet.
    if let Some(id) = numeric_part_stem(name, "/xl/drawings/drawing", ".xml") {
        return !worksheet_ids.contains(&id);
    }
    if let Some(id) = numeric_part_stem(name, "/xl/drawings/vmlDrawing", ".vml") {
        return !worksheet_ids.contains(&id);
    }
    // /xl/comments{id}.xml and its threaded counterpart — per worksheet.
    if let Some(id) = numeric_part_stem(name, "/xl/comments", ".xml") {
        return !worksheet_ids.contains(&id);
    }
    if let Some(id) = numeric_part_stem(name, "/xl/threadedComments/threadedComment", ".xml") {
        return !worksheet_ids.contains(&id);
    }
    // /xl/media/image{id}_{m}.{ext} — per worksheet.
    if let Some(rest) = name.strip_prefix("/xl/media/image") {
        let digits = rest.bytes().take_while(u8::is_ascii_digit).count();
        if digits > 0
            && rest[digits..].starts_with('_')
            && let Ok(id) = rest[..digits].parse::<u32>()
        {
            return !worksheet_ids.contains(&id);
        }
    }
    false
}

fn next_active_x_part_uri(
    package: &OpcPackage,
    directory: &str,
    stem: &str,
    extension: &str,
) -> SheetResult<PackURI> {
    for suffix in 0..=100_000usize {
        let uri = PackURI::new(format!("{directory}/{stem}_{suffix}.{extension}"))?;
        if package.validate_new_part_name(&uri).is_ok() {
            return Ok(uri);
        }
    }
    Err("Unable to allocate an ActiveX part name".into())
}

/// Concrete implementation of a Workbook for Excel files.
#[derive(Debug)]
pub struct Workbook {
    /// The underlying OPC package
    package: OpcPackage,
    /// Actual main-workbook part location resolved from the package relationship.
    workbook_uri: PackURI,
    /// Cached worksheet information
    worksheets: Vec<WorksheetInfo>,
    /// Cached worksheet names for zero-copy returns
    worksheet_names: Vec<String>,
    /// Active worksheet index (0-based)
    active_sheet_index: usize,
    /// Shared strings table for efficient string storage
    shared_strings: SharedStrings,
    /// Styles information
    styles: Styles,
    /// Mutable workbook data for writing (cached)
    mutable_data: Option<MutableWorkbookData>,
    /// Document properties (metadata)
    properties: DocumentProperties,
    /// Whether the workbook uses the 1904 date system
    is_1904_date_system: bool,
    /// Effective workbook formula calculation policy.
    calculation_properties: Option<WorkbookCalculationProperties>,
    /// Inert workbook calculation order from `calcChain.xml`.
    calculation_chain: Option<CalculationChain>,
    /// Namespace family of the cached calculation chain, retained on writer materialization.
    calculation_chain_conformance: Option<CalculationChainConformance>,
    external_links: Vec<ExternalLinkEntry>,
    defined_names: Vec<NamedRange>,
    worksheet_protection_mutations: HashMap<usize, WorksheetProtectionMetadata>,
    worksheet_data_validation_mutations: HashMap<usize, Vec<DataValidationCollection>>,
    worksheet_web_extension_binding_mutations: HashMap<usize, Vec<WorksheetWebExtensionBinding>>,
}

fn patch_workbook_external_references(
    xml: &[u8],
    relationship_ids: &[String],
    conformance: ExternalLinkConformance,
) -> SheetResult<Vec<u8>> {
    use quick_xml::Reader;
    use quick_xml::events::Event;

    if xml.len() > 64 * 1024 * 1024 {
        return Err("workbook XML exceeds the external-link mutation limit".into());
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_name = Vec::new();
    let mut root_prefix = String::new();
    let mut relationship_prefix_bound = false;
    let mut existing_start = None;
    let mut existing_end = None;
    let mut insertion = None;
    let mut active_external_depth = None;
    loop {
        let before = reader.buffer_position() as usize;
        match reader.read_event()? {
            Event::Start(element) => {
                let name = element.name().as_ref().to_vec();
                let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(&name);
                if depth == 0 {
                    root_name = name.clone();
                    root_prefix = name
                        .split(|byte| *byte == b':')
                        .next()
                        .filter(|_| name.contains(&b':'))
                        .map(|value| String::from_utf8_lossy(value).into_owned())
                        .unwrap_or_default();
                    for attribute in element.attributes().with_checks(false) {
                        let attribute = attribute?;
                        if attribute.key.as_ref() == b"xmlns:r" {
                            relationship_prefix_bound = true;
                        }
                    }
                } else if depth == 1 {
                    if local == b"externalReferences" {
                        if existing_start.replace(before).is_some() {
                            return Err("workbook has multiple externalReferences elements".into());
                        }
                        active_external_depth = Some(depth);
                    } else if insertion.is_none()
                        && matches!(
                            local,
                            b"definedNames"
                                | b"calcPr"
                                | b"oleSize"
                                | b"customWorkbookViews"
                                | b"pivotCaches"
                                | b"smartTagPr"
                                | b"smartTagTypes"
                                | b"webPublishing"
                                | b"fileRecoveryPr"
                                | b"webPublishObjects"
                                | b"extLst"
                        )
                    {
                        insertion = Some(before);
                    }
                }
                depth = depth.checked_add(1).ok_or("workbook XML depth overflow")?;
            },
            Event::Empty(element) => {
                let name = element.name().as_ref().to_vec();
                let local = name.rsplit(|byte| *byte == b':').next().unwrap_or(&name);
                if depth == 1 && local == b"externalReferences" {
                    return Err("workbook externalReferences cannot be empty".into());
                }
                if depth == 1
                    && insertion.is_none()
                    && matches!(
                        local,
                        b"definedNames"
                            | b"calcPr"
                            | b"oleSize"
                            | b"customWorkbookViews"
                            | b"pivotCaches"
                            | b"smartTagPr"
                            | b"smartTagTypes"
                            | b"webPublishing"
                            | b"fileRecoveryPr"
                            | b"webPublishObjects"
                            | b"extLst"
                    )
                {
                    insertion = Some(before);
                }
            },
            Event::End(element) => {
                depth = depth.checked_sub(1).ok_or("invalid workbook XML depth")?;
                if active_external_depth == Some(depth) {
                    existing_end = Some(reader.buffer_position() as usize);
                    active_external_depth = None;
                }
                if depth == 0 {
                    if element.name().as_ref() != root_name.as_slice() {
                        return Err("workbook root element is unbalanced".into());
                    }
                    insertion.get_or_insert(before);
                    break;
                }
            },
            Event::Eof => return Err("workbook XML ended before the root closed".into()),
            _ => {},
        }
    }

    let mut fragment = Vec::new();
    if !relationship_ids.is_empty() {
        let prefix = if root_prefix.is_empty() {
            String::new()
        } else {
            format!("{}:", root_prefix)
        };
        let relationship_namespace = match conformance {
            ExternalLinkConformance::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            ExternalLinkConformance::Strict => {
                "http://purl.oclc.org/ooxml/officeDocument/relationships"
            },
        };
        let mut value = format!("<{prefix}externalReferences");
        if !relationship_prefix_bound {
            value.push_str(" xmlns:r=\"");
            value.push_str(relationship_namespace);
            value.push('"');
        }
        value.push('>');
        for relationship_id in relationship_ids {
            if relationship_id.is_empty()
                || relationship_id.len() > 1024
                || relationship_id.chars().any(char::is_control)
            {
                return Err("invalid external-link relationship ID".into());
            }
            value.push('<');
            value.push_str(&prefix);
            value.push_str("externalReference r:id=\"");
            value.push_str(&litchi_core::xml::escape_xml(relationship_id));
            value.push_str("\"/>");
        }
        value.push_str("</");
        value.push_str(&prefix);
        value.push_str("externalReferences>");
        fragment = value.into_bytes();
    }
    let (start, end) = match (existing_start, existing_end) {
        (Some(start), Some(end)) => (start, end),
        (None, None) => {
            let position = insertion.ok_or("workbook insertion point is missing")?;
            (position, position)
        },
        _ => return Err("workbook externalReferences element is unbalanced".into()),
    };
    let new_len = xml
        .len()
        .checked_sub(end - start)
        .and_then(|length| length.checked_add(fragment.len()))
        .ok_or("workbook externalReferences size overflow")?;
    if new_len > 64 * 1024 * 1024 {
        return Err("mutated workbook XML exceeds 64 MiB".into());
    }
    let mut output = Vec::with_capacity(new_len);
    output.extend_from_slice(&xml[..start]);
    output.extend_from_slice(&fragment);
    output.extend_from_slice(&xml[end..]);
    Ok(output)
}

impl Workbook {
    /// Discover inert embedded-object and embedded-package relationships.
    pub fn embedded_parts(&self) -> crate::error::Result<Vec<crate::EmbeddedPart<'_>>> {
        crate::embedded_object::discover_embedded_parts(&self.package)
    }

    /// Discover the attached MS-OFFMACRO2 VBA project without inspecting its payload.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// type. It does not inspect, parse, decompress, or execute the binary
    /// VBA project bytes.
    pub fn vba_project(&self) -> crate::error::Result<Option<VbaProject>> {
        let workbook = self.package.get_part(&self.workbook_uri)?;
        discover_vba_project(&self.package, workbook)
    }

    /// The theme part of this workbook (ECMA-376 DrawingML theme), when
    /// present. Theme-indexed colors in fonts, fills, borders, tab colors,
    /// and charts resolve against the returned color scheme.
    pub fn theme(&self) -> crate::error::Result<Option<crate::xlsx::theme::XlsxTheme>> {
        use litchi_opc::constants::relationship_type as rt;

        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let declared = workbook_part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rt::THEME)
            .map(|relationship| relationship.target_partname())
            .transpose()?;
        let (uri, well_known) = match declared {
            Some(uri) => (uri, false),
            None => (
                PackURI::new("/xl/theme/theme1.xml")
                    .map_err(crate::error::OoxmlError::InvalidFormat)?,
                true,
            ),
        };
        let part = match self.package.get_part(&uri) {
            Ok(part) => part,
            Err(error) if well_known => {
                let _ = error;
                return Ok(None);
            },
            Err(error) => return Err(error.into()),
        };
        let xml = std::str::from_utf8(part.blob())
            .map_err(|error| crate::error::OoxmlError::InvalidFormat(error.to_string()))?;
        crate::xlsx::theme::XlsxTheme::parse(xml).map(Some)
    }

    /// Attach a cache-free, inert MS-OVBA project and convert this package to XLSM/XLTM.
    pub fn set_vba_project(
        &mut self,
        project: &crate::vba::VbaProjectBinary,
    ) -> crate::error::Result<VbaProject> {
        let payload = project
            .to_cfb_bytes()
            .map_err(|error| crate::error::OoxmlError::InvalidFormat(error.to_string()))?;
        self.set_vba_project_bytes(payload, &crate::vba::VbaLimits::default())
    }

    /// Attach an existing, validated `vbaProject.bin` payload without executing it.
    pub fn set_vba_project_bytes(
        &mut self,
        payload: Vec<u8>,
        limits: &crate::vba::VbaLimits,
    ) -> crate::error::Result<VbaProject> {
        store_workbook_vba_project(&mut self.package, &self.workbook_uri, payload, limits)
    }

    /// Remove the VBA project graph and convert XLSM/XLTM content types back to XLSX/XLTX.
    pub fn remove_vba_project(&mut self) -> crate::error::Result<bool> {
        remove_workbook_vba_project(&mut self.package, &self.workbook_uri)
    }

    /// Load persisted Office Add-in task-pane metadata without activating add-ins.
    pub fn web_extension_task_panes(&self) -> crate::error::Result<Option<WebExtensionTaskPanes>> {
        load_web_extension_task_panes(&self.package)
    }

    /// Store inert persisted Office Add-in task panes and snapshot resources.
    pub fn set_web_extension_task_panes(
        &mut self,
        task_panes: &WebExtensionTaskPanes,
        conformance: OoxmlConformance,
    ) -> crate::error::Result<()> {
        store_web_extension_task_panes(&mut self.package, task_panes, conformance)
    }

    /// Remove persisted Office Add-in task panes and unreferenced resources.
    pub fn remove_web_extension_task_panes(&mut self) -> crate::error::Result<bool> {
        remove_web_extension_task_panes(&mut self.package)
    }

    /// Load all package-level RibbonX customizations without invoking callbacks.
    pub fn ribbon_customizations(&self) -> crate::error::Result<Vec<RibbonCustomization>> {
        load_ribbon_customizations(&self.package)
    }

    /// Load the effective package-level RibbonX customization without invoking callbacks.
    pub fn ribbon_customization(&self) -> crate::error::Result<Option<RibbonCustomization>> {
        load_ribbon_customization(&self.package)
    }

    /// Store opaque RibbonX XML without interpreting or invoking callbacks.
    pub fn set_ribbon_customization(
        &mut self,
        version: RibbonCustomizationVersion,
        xml: &[u8],
    ) -> crate::error::Result<RibbonCustomization> {
        let customization = store_ribbon_customization(&mut self.package, version, xml)?;
        let _ = self.package.clear_digital_signatures();
        Ok(customization)
    }

    /// Create a new empty workbook.
    ///
    /// Creates a minimal valid Excel workbook with one default worksheet.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut workbook = Workbook::create()?;
    /// // Add data to worksheets...
    /// workbook.save("new_workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn create() -> SheetResult<Self> {
        use crate::xlsx::template;
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::BlobPart;

        let mut package = OpcPackage::new();

        // Create workbook.xml
        let workbook_uri = PackURI::new("/xl/workbook.xml")?;
        let workbook_part = BlobPart::new(
            workbook_uri.clone(),
            ct::SML_SHEET_MAIN.to_string(),
            template::default_workbook_xml().as_bytes().to_vec(),
        );
        // Use relative path for package-level relationship
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));

        // Create worksheet
        let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml")?;
        let worksheet_part = BlobPart::new(
            worksheet_uri,
            ct::SML_WORKSHEET.to_string(),
            template::default_worksheet_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to(
                "worksheets/sheet1.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            );
        }
        package.add_part(Box::new(worksheet_part));

        // Create styles.xml
        let styles_uri = PackURI::new("/xl/styles.xml")?;
        let styles_part = BlobPart::new(
            styles_uri,
            ct::SML_STYLES.to_string(),
            template::default_styles_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to("styles.xml", rt::STYLES);
        }
        package.add_part(Box::new(styles_part));

        // Create sharedStrings.xml
        let shared_strings_uri = PackURI::new("/xl/sharedStrings.xml")?;
        let shared_strings_part = BlobPart::new(
            shared_strings_uri,
            ct::SML_SHARED_STRINGS.to_string(),
            template::default_shared_strings_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to(
                "sharedStrings.xml",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            );
        }
        package.add_part(Box::new(shared_strings_part));

        // Create theme
        let theme_uri = PackURI::new("/xl/theme/theme1.xml")?;
        let theme_part = BlobPart::new(
            theme_uri,
            ct::OFC_THEME.to_string(),
            template::default_theme_xml().as_bytes().to_vec(),
        );
        if let Ok(wb_part) = package.get_part_mut(&workbook_uri) {
            wb_part.relate_to("theme/theme1.xml", rt::THEME);
        }
        package.add_part(Box::new(theme_part));

        // Create core.xml
        let core_props_uri = PackURI::new("/docProps/core.xml")?;
        let core_props_part = BlobPart::new(
            core_props_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            template::default_core_props_xml().as_bytes().to_vec(),
        );
        package.relate_to("docProps/core.xml", rt::CORE_PROPERTIES);
        package.add_part(Box::new(core_props_part));

        // Create app.xml
        let app_props_uri = PackURI::new("/docProps/app.xml")?;
        let app_props_part = BlobPart::new(
            app_props_uri,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            template::default_app_props_xml().as_bytes().to_vec(),
        );
        package.relate_to("docProps/app.xml", rt::EXTENDED_PROPERTIES);
        package.add_part(Box::new(app_props_part));

        let mut workbook = Self::new(package)?;
        workbook.mutable_data = Some(MutableWorkbookData::new());
        Ok(workbook)
    }

    /// Create a new workbook from an OPC package.
    pub fn new(package: OpcPackage) -> SheetResult<Self> {
        use litchi_opc::constants::content_type as ct;

        let workbook_part = package.main_document_part()?;
        if !matches!(
            workbook_part.content_type(),
            ct::SML_SHEET_MAIN
                | ct::SML_TEMPLATE_MAIN
                | ct::SML_SHEET_MACRO_MAIN
                | ct::SML_TEMPLATE_MACRO_MAIN
        ) {
            return Err(format!(
                "main document part '{}' is not an XML workbook",
                workbook_part.partname()
            )
            .into());
        }
        let workbook_uri = workbook_part.partname().clone();
        let mut workbook = Workbook {
            package,
            workbook_uri,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            active_sheet_index: 0,
            shared_strings: SharedStrings::new(),
            styles: Styles::new(),
            mutable_data: None,
            properties: DocumentProperties::new(),
            is_1904_date_system: false,
            calculation_properties: None,
            calculation_chain: None,
            calculation_chain_conformance: None,
            external_links: Vec::new(),
            defined_names: Vec::new(),
            worksheet_protection_mutations: HashMap::new(),
            worksheet_data_validation_mutations: HashMap::new(),
            worksheet_web_extension_binding_mutations: HashMap::new(),
        };

        workbook.load_workbook_info()?;
        if let Some((chain, conformance)) =
            load_calculation_chain(&workbook.package, &workbook.workbook_uri)?
        {
            workbook.calculation_chain = Some(chain);
            workbook.calculation_chain_conformance = Some(conformance);
        }
        workbook.load_external_links()?;
        workbook.load_shared_strings()?;
        workbook.load_styles()?;

        Ok(workbook)
    }

    /// Load workbook information from workbook.xml
    fn load_workbook_info(&mut self) -> SheetResult<()> {
        let workbook_part = self.package.get_part(&self.workbook_uri)?;

        // Parse the workbook XML to extract sheet information
        let content = std::str::from_utf8(workbook_part.blob())?;

        // Extract sheets from workbook.xml
        let mut details = workbook_parser::parse_workbook_details(content)?;
        let calculation_properties = parse_workbook_calculation_properties(content.as_bytes())?;
        Self::apply_defined_name_print_settings(&details.defined_names, &mut details.sheets);

        // Cache worksheet names for zero-copy returns
        self.worksheet_names = details.sheets.iter().map(|ws| ws.name.clone()).collect();
        self.worksheets = details.sheets;
        self.active_sheet_index = details.active_sheet_index;
        self.is_1904_date_system = details.uses_1904_date_system;
        self.calculation_properties = calculation_properties;
        self.defined_names = details.defined_names;

        Ok(())
    }

    fn load_external_links(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};
        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let content = std::str::from_utf8(workbook_part.blob())?;
        let details = workbook_parser::parse_workbook_details(content)?;
        let mut links = Vec::with_capacity(details.external_reference_ids.len());
        for (offset, relationship_id) in details.external_reference_ids.into_iter().enumerate() {
            let relationship = workbook_part.rels().get(&relationship_id).ok_or_else(|| {
                format!("workbook external reference '{relationship_id}' has no relationship")
            })?;
            if relationship.is_external()
                || !matches!(
                    relationship.reltype(),
                    rt::EXTERNAL_LINK | rt::STRICT_EXTERNAL_LINK
                )
            {
                return Err(format!(
                    "workbook external reference '{relationship_id}' has an invalid relationship"
                )
                .into());
            }
            let uri = relationship.target_partname()?;
            let part = self.package.get_part(&uri)?;
            if part.content_type() != ct::SML_EXTERNAL_LINK {
                return Err(format!(
                    "external-link part '{uri}' has invalid content type '{}', expected '{}'",
                    part.content_type(),
                    ct::SML_EXTERNAL_LINK
                )
                .into());
            }
            let index = u32::try_from(offset + 1).map_err(|_| "external-link index overflow")?;
            links.push(load_external_link(part, relationship_id, index)?);
        }
        self.external_links = links;
        Ok(())
    }

    pub fn external_links(&self) -> &[ExternalLinkEntry] {
        &self.external_links
    }

    /// Effective workbook formula calculation policy, when `calcPr` is present.
    pub fn calculation_properties(&self) -> Option<&WorkbookCalculationProperties> {
        self.calculation_properties.as_ref()
    }

    /// Return the optional inert calculation-chain metadata.
    pub fn calculation_chain(&self) -> Option<&CalculationChain> {
        self.calculation_chain.as_ref()
    }

    /// Return the XML and relationship namespace family of the calculation chain.
    pub fn calculation_chain_conformance(&self) -> Option<CalculationChainConformance> {
        self.calculation_chain_conformance
    }

    /// Replace the workbook's inert calculation-chain metadata.
    ///
    /// This serializes the caller-authored order only. It neither evaluates
    /// formulas nor rebuilds dependencies from worksheet formulas.
    pub fn set_calculation_chain(
        &mut self,
        chain: CalculationChain,
        conformance: CalculationChainConformance,
    ) -> SheetResult<()> {
        store_calculation_chain(&mut self.package, &chain, conformance)?;
        self.calculation_chain = Some(chain);
        self.calculation_chain_conformance = Some(conformance);
        Ok(())
    }

    /// Remove the optional calculation chain without changing worksheet formulas.
    pub fn remove_calculation_chain(&mut self) -> SheetResult<bool> {
        let removed = remove_calculation_chain(&mut self.package)?;
        self.calculation_chain = None;
        self.calculation_chain_conformance = None;
        Ok(removed)
    }

    /// Load inert Custom XML Maps metadata and its namespace family.
    ///
    /// Inline schemas and data bindings are retained as metadata only. This
    /// never resolves schema locations, opens files, or imports/exports mapped
    /// worksheet data.
    pub fn xml_maps(&self) -> SheetResult<Option<(XmlMapInfo, XmlMapConformance)>> {
        load_xml_maps(&self.package)
    }

    /// Replace the workbook's inert Custom XML Maps metadata.
    ///
    /// This writes only the caller-provided MapInfo part. It does not apply a
    /// mapping, resolve a schema, or interact with bound external content.
    pub fn set_xml_maps(
        &mut self,
        value: &XmlMapInfo,
        conformance: XmlMapConformance,
    ) -> SheetResult<()> {
        store_xml_maps(&mut self.package, value, conformance)
    }

    /// Remove Custom XML Maps metadata without changing worksheet cell data.
    pub fn remove_xml_maps(&mut self) -> SheetResult<bool> {
        remove_xml_maps(&mut self.package)
    }

    /// Load inert volatile-dependencies metadata and its namespace family.
    ///
    /// This never contacts RTD servers, opens OLAP connections, or evaluates
    /// workbook formulas.
    pub fn volatile_dependencies(
        &self,
    ) -> SheetResult<Option<(VolatileDependencies, VolatileDependenciesConformance)>> {
        load_volatile_dependencies(&self.package)
    }

    /// Replace the workbook's inert volatile-dependencies metadata.
    ///
    /// The caller-provided records are serialized without RTD, cube, or formula
    /// evaluation work.
    pub fn set_volatile_dependencies(
        &mut self,
        value: &VolatileDependencies,
        conformance: VolatileDependenciesConformance,
    ) -> SheetResult<()> {
        store_volatile_dependencies(&mut self.package, value, conformance)
    }

    /// Remove volatile-dependencies metadata without recalculating formulas.
    pub fn remove_volatile_dependencies(&mut self) -> SheetResult<bool> {
        remove_volatile_dependencies(&mut self.package)
    }

    /// Load the optional Named Sheet Views attached to one worksheet.
    ///
    /// These are stored sort/filter settings only. Reading them does not apply
    /// filters, reorder cells, evaluate formulas, or fetch external data.
    pub fn named_sheet_views(&self, index: usize) -> SheetResult<Option<NamedSheetViews>> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        load_worksheet_named_sheet_views(&self.package, &uri).map_err(Into::into)
    }

    /// Store caller-authored Named Sheet Views for one worksheet.
    ///
    /// The supplied metadata is serialized without applying its sort or filter
    /// settings to worksheet cells.
    pub fn set_named_sheet_views(
        &mut self,
        index: usize,
        value: &NamedSheetViews,
    ) -> SheetResult<()> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        store_worksheet_named_sheet_views(&mut self.package, &uri, value).map_err(Into::into)
    }

    /// Remove the optional Named Sheet Views from one worksheet without
    /// changing the worksheet's ordinary active view or cell data.
    pub fn remove_named_sheet_views(&mut self, index: usize) -> SheetResult<bool> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        remove_worksheet_named_sheet_views(&mut self.package, &uri).map_err(Into::into)
    }

    /// Load one worksheet's ActiveX descriptors and opaque persistence bytes.
    ///
    /// No control, callback, or embedded binary is instantiated or executed.
    pub fn worksheet_active_x_controls(&self, index: usize) -> SheetResult<ActiveXControlSet> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        load_worksheet_active_x(&self.package, &uri).map_err(Into::into)
    }

    /// Store a complete inert ActiveX graph on a worksheet without controls.
    pub fn store_worksheet_active_x_controls(
        &mut self,
        index: usize,
        value: &ActiveXControlSet,
    ) -> SheetResult<()> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        store_worksheet_active_x(&mut self.package, &uri, value).map_err(Into::into)
    }

    /// Atomically replace one worksheet's complete inert ActiveX graph.
    pub fn replace_worksheet_active_x_controls(
        &mut self,
        index: usize,
        value: &ActiveXControlSet,
    ) -> SheetResult<()> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        replace_worksheet_active_x(&mut self.package, &uri, value).map_err(Into::into)
    }

    /// Remove one worksheet's ActiveX graph without activating its payloads.
    pub fn remove_worksheet_active_x_controls(&mut self, index: usize) -> SheetResult<bool> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        remove_worksheet_active_x(&mut self.package, &uri).map_err(Into::into)
    }

    /// Return passive `workbookProtection` metadata from the current `workbook.xml` part.
    ///
    /// Password verifier values remain opaque: this method never accepts or
    /// checks a password, and it does not enforce the requested locks.
    pub fn workbook_protection_metadata(&self) -> SheetResult<Option<WorkbookProtectionMetadata>> {
        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        parse_workbook_protection(workbook_part.blob()).map_err(Into::into)
    }

    pub fn external_link(&self, one_based_index: u32) -> Option<&ExternalLinkEntry> {
        one_based_index
            .checked_sub(1)
            .and_then(|index| self.external_links.get(index as usize))
    }

    /// Find an external link by its workbook relationship ID.
    pub fn find_external_link_by_relationship_id(
        &self,
        relationship_id: &str,
    ) -> Option<&ExternalLinkEntry> {
        self.external_links
            .iter()
            .find(|entry| entry.relationship_id == relationship_id)
    }

    /// Find inert workbook/OLE links whose relationship target matches exactly.
    pub fn find_external_links_by_target(&self, target: &str) -> Vec<&ExternalLinkEntry> {
        self.external_links
            .iter()
            .filter(|entry| match &entry.kind {
                ExternalLinkKind::Workbook(link) => link.target.target == target,
                ExternalLinkKind::Ole(link) => link.target.target == target,
                ExternalLinkKind::Dde(_) => false,
            })
            .collect()
    }

    /// Add a fully typed external-workbook, DDE, or OLE link without opening its target.
    pub fn add_external_link(&mut self, kind: ExternalLinkKind) -> SheetResult<u32> {
        let conformance = self.external_link_conformance()?;
        self.add_external_link_with_conformance(kind, conformance)
    }

    /// Add an inert external link using the requested strict/transitional namespaces.
    pub fn add_external_link_with_conformance(
        &mut self,
        kind: ExternalLinkKind,
        conformance: ExternalLinkConformance,
    ) -> SheetResult<u32> {
        if self.external_links.len() >= 4096 {
            return Err("workbook external-link limit exceeded".into());
        }
        let mut selected = None;
        for suffix in 1..=4097u32 {
            let uri = PackURI::new(format!("/xl/externalLinks/externalLink{suffix}.xml"))?;
            if self.package.get_part(&uri).is_err() {
                selected = Some((suffix, uri));
                break;
            }
        }
        let (suffix, part_uri) = selected.ok_or("no free external-link part name")?;
        let part = build_external_link_part_with_conformance(part_uri.clone(), &kind, conformance)?;
        let target = format!("externalLinks/externalLink{suffix}.xml");
        let relationship_id = self.next_workbook_relationship_id("rIdExternalLink")?;
        let mut ids = self
            .external_links
            .iter()
            .map(|entry| entry.relationship_id.clone())
            .collect::<Vec<_>>();
        ids.push(relationship_id.clone());
        let original = self.package.get_part(&self.workbook_uri)?.blob().to_vec();
        let replacement = patch_workbook_external_references(&original, &ids, conformance)?;
        workbook_parser::parse_workbook_details(std::str::from_utf8(&replacement)?)?;
        self.package.try_add_part(Box::new(part))?;
        let workbook = self.package.get_part_mut(&self.workbook_uri)?;
        workbook.rels_mut().add_relationship(
            conformance.external_link_relationship().to_string(),
            target,
            relationship_id.clone(),
            false,
        );
        workbook.set_blob(replacement);
        let index = u32::try_from(self.external_links.len() + 1)
            .map_err(|_| "external-link index overflow")?;
        self.external_links.push(ExternalLinkEntry {
            index,
            relationship_id,
            part_uri,
            kind,
        });
        let _ = self.package.clear_digital_signatures();
        Ok(index)
    }

    /// Replace a typed external link while preserving its workbook relationship and part URI.
    pub fn replace_external_link(
        &mut self,
        one_based_index: u32,
        kind: ExternalLinkKind,
    ) -> SheetResult<()> {
        let offset = one_based_index
            .checked_sub(1)
            .ok_or("external-link indices are one-based")? as usize;
        let entry = self
            .external_links
            .get(offset)
            .ok_or("external-link index is out of bounds")?;
        let part_uri = entry.part_uri.clone();
        let conformance = self.external_link_conformance()?;
        let part = build_external_link_part_with_conformance(part_uri, &kind, conformance)?;
        load_external_link(&part, entry.relationship_id.clone(), one_based_index)?;
        self.package.add_part(Box::new(part));
        self.external_links[offset].kind = kind;
        let _ = self.package.clear_digital_signatures();
        Ok(())
    }

    /// Alias for index-stable replacement.
    pub fn update_external_link(
        &mut self,
        one_based_index: u32,
        kind: ExternalLinkKind,
    ) -> SheetResult<()> {
        self.replace_external_link(one_based_index, kind)
    }

    /// Remove one external link when doing so cannot reinterpret a formula index.
    pub fn remove_external_link(&mut self, one_based_index: u32) -> SheetResult<ExternalLinkEntry> {
        let offset = one_based_index
            .checked_sub(1)
            .ok_or("external-link indices are one-based")? as usize;
        if offset >= self.external_links.len() {
            return Err("external-link index is out of bounds".into());
        }
        let affected = (one_based_index..=self.external_links.len() as u32).collect::<Vec<_>>();
        self.ensure_external_formula_indices_unused(&affected)?;
        let conformance = self.external_link_conformance()?;
        let ids = self
            .external_links
            .iter()
            .enumerate()
            .filter(|(index, _)| *index != offset)
            .map(|(_, entry)| entry.relationship_id.clone())
            .collect::<Vec<_>>();
        let original = self.package.get_part(&self.workbook_uri)?.blob().to_vec();
        let replacement = patch_workbook_external_references(&original, &ids, conformance)?;
        workbook_parser::parse_workbook_details(std::str::from_utf8(&replacement)?)?;

        let removed = self.external_links.remove(offset);
        let workbook = self.package.get_part_mut(&self.workbook_uri)?;
        workbook.rels_mut().remove(&removed.relationship_id);
        workbook.set_blob(replacement);
        for (index, entry) in self.external_links.iter_mut().enumerate() {
            entry.index = u32::try_from(index + 1).map_err(|_| "external-link index overflow")?;
        }
        if !self.package_part_is_referenced(&removed.part_uri) {
            self.package.remove_part(&removed.part_uri);
        }
        let _ = self.package.clear_digital_signatures();
        Ok(removed)
    }

    /// Reorder links by their current one-based indices.
    ///
    /// The operation is rejected when a moved index appears in package formula text; this
    /// prevents a reorder from silently changing the workbook a formula points at.
    pub fn reorder_external_links(&mut self, order: &[u32]) -> SheetResult<()> {
        if order.len() != self.external_links.len() {
            return Err("external-link reorder must contain every link exactly once".into());
        }
        let expected = (1..=self.external_links.len() as u32).collect::<HashSet<_>>();
        let actual = order.iter().copied().collect::<HashSet<_>>();
        if actual != expected || actual.len() != order.len() {
            return Err("external-link reorder is not a permutation".into());
        }
        let moved = order
            .iter()
            .enumerate()
            .filter_map(|(new, old)| (*old != new as u32 + 1).then_some(*old))
            .collect::<Vec<_>>();
        self.ensure_external_formula_indices_unused(&moved)?;
        let reordered = order
            .iter()
            .map(|index| self.external_links[(*index - 1) as usize].clone())
            .collect::<Vec<_>>();
        let ids = reordered
            .iter()
            .map(|entry| entry.relationship_id.clone())
            .collect::<Vec<_>>();
        let conformance = self.external_link_conformance()?;
        let original = self.package.get_part(&self.workbook_uri)?.blob().to_vec();
        let replacement = patch_workbook_external_references(&original, &ids, conformance)?;
        workbook_parser::parse_workbook_details(std::str::from_utf8(&replacement)?)?;
        self.package
            .get_part_mut(&self.workbook_uri)?
            .set_blob(replacement);
        self.external_links = reordered;
        for (index, entry) in self.external_links.iter_mut().enumerate() {
            entry.index = u32::try_from(index + 1).map_err(|_| "external-link index overflow")?;
        }
        let _ = self.package.clear_digital_signatures();
        Ok(())
    }

    fn external_link_conformance(&self) -> SheetResult<ExternalLinkConformance> {
        let xml = self.package.get_part(&self.workbook_uri)?.blob();
        if xml
            .windows(b"http://purl.oclc.org/ooxml/spreadsheetml/main".len())
            .any(|window| window == b"http://purl.oclc.org/ooxml/spreadsheetml/main")
        {
            Ok(ExternalLinkConformance::Strict)
        } else {
            Ok(ExternalLinkConformance::Transitional)
        }
    }

    fn next_workbook_relationship_id(&self, prefix: &str) -> SheetResult<String> {
        let relationships = self.package.get_part(&self.workbook_uri)?.rels();
        for suffix in 1..=65_537u32 {
            let candidate = format!("{prefix}{suffix}");
            if relationships.get(&candidate).is_none() {
                return Ok(candidate);
            }
        }
        Err("no free workbook relationship ID".into())
    }

    fn ensure_external_formula_indices_unused(&self, indices: &[u32]) -> SheetResult<()> {
        if indices.is_empty() {
            return Ok(());
        }
        let needles = indices
            .iter()
            .map(|index| format!("[{index}]"))
            .collect::<Vec<_>>();
        if self.package.iter_parts().any(|part| {
            let Ok(text) = std::str::from_utf8(part.blob()) else {
                return false;
            };
            needles.iter().any(|needle| text.contains(needle))
        }) {
            return Err(
                "external-link operation would change an index referenced by formula metadata"
                    .into(),
            );
        }
        Ok(())
    }

    fn package_part_is_referenced(&self, part_uri: &PackURI) -> bool {
        self.package.iter_parts().any(|part| {
            part.rels().iter().any(|relationship| {
                !relationship.is_external()
                    && relationship
                        .target_partname()
                        .is_ok_and(|target| target == *part_uri)
            })
        }) || self.package.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|target| target == *part_uri)
        })
    }

    /// Load shared strings from xl/sharedStrings.xml
    fn load_shared_strings(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        let Some(shared_strings_uri) = self.related_workbook_part_uri(
            &[rt::SHARED_STRINGS, rt::STRICT_SHARED_STRINGS],
            "shared strings",
        )?
        else {
            return Ok(());
        };
        let shared_strings_part = self.package.get_part(&shared_strings_uri)?;
        Self::require_part_content_type(
            &shared_strings_uri,
            shared_strings_part.content_type(),
            ct::SML_SHARED_STRINGS,
        )?;
        let content = std::str::from_utf8(shared_strings_part.blob())?;
        self.shared_strings = SharedStrings::parse(content)?;

        Ok(())
    }

    /// Load styles from xl/styles.xml
    fn load_styles(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::{content_type as ct, relationship_type as rt};

        let Some(styles_uri) =
            self.related_workbook_part_uri(&[rt::STYLES, rt::STRICT_STYLES], "styles")?
        else {
            return Ok(());
        };
        let styles_part = self.package.get_part(&styles_uri)?;
        Self::require_part_content_type(&styles_uri, styles_part.content_type(), ct::SML_STYLES)?;
        let content = std::str::from_utf8(styles_part.blob())?;
        self.styles = Styles::parse(content)
            .map_err(|e| -> Box<dyn std::error::Error + Send + Sync> { Box::new(e) })?;
        Ok(())
    }

    fn related_workbook_part_uri(
        &self,
        relationship_types: &[&str],
        description: &str,
    ) -> SheetResult<Option<PackURI>> {
        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let mut matching = workbook_part
            .rels()
            .iter()
            .filter(|relationship| relationship_types.contains(&relationship.reltype()));
        let Some(relationship) = matching.next() else {
            return Ok(None);
        };
        if matching.next().is_some() {
            return Err(format!("workbook has multiple {description} relationships").into());
        }
        if relationship.is_external() {
            return Err(format!("workbook {description} relationship cannot be external").into());
        }
        Ok(Some(relationship.target_partname()?))
    }

    fn require_part_content_type(uri: &PackURI, actual: &str, expected: &str) -> SheetResult<()> {
        if actual != expected {
            return Err(
                format!("part '{uri}' has content type '{actual}', expected '{expected}'").into(),
            );
        }
        Ok(())
    }

    /// Apply sheet-scoped print-area and print-title defined names.
    fn apply_defined_name_print_settings(
        defined_names: &[NamedRange],
        worksheets: &mut [WorksheetInfo],
    ) {
        for defined_name in defined_names {
            let Some(sheet_idx) = defined_name
                .local_sheet_id
                .and_then(|index| usize::try_from(index).ok())
            else {
                continue;
            };
            let worksheet = &mut worksheets[sheet_idx];
            if defined_name.name == "_xlnm.Print_Area" {
                worksheet.print_area = Self::parse_print_area(&defined_name.reference);
            } else if defined_name.name == "_xlnm.Print_Titles" {
                let (rows, columns) = Self::parse_print_titles(&defined_name.reference);
                worksheet.repeating_rows = rows;
                worksheet.repeating_columns = columns;
            }
        }
    }

    /// Return all workbook- and sheet-scoped defined names as inert metadata.
    pub fn defined_names(&self) -> &[NamedRange] {
        &self.defined_names
    }

    /// Parse the print area reference from a defined name value.
    ///
    /// Values are typically of the form `'Sheet Name'!A1:D20` or a comma-
    /// separated list of such references. We return the range part for the
    /// first entry (e.g., `A1:D20`).
    fn parse_print_area(value: &str) -> Option<String> {
        let first = Self::split_defined_name_areas(value).into_iter().next()?;
        let bang = first.rfind('!')?;
        let range = first[bang + 1..].trim();
        Self::is_valid_print_cell_range(range).then(|| range.to_string())
    }

    /// Parse repeating rows/columns from a _xlnm.Print_Titles defined name
    /// value. Returns (rows, columns) as raw range strings (e.g., "$1:$1",
    /// "$A:$B").
    fn parse_print_titles(value: &str) -> (Option<String>, Option<String>) {
        let mut rows: Option<String> = None;
        let mut cols: Option<String> = None;

        for part in Self::split_defined_name_areas(value) {
            let bang = match part.rfind('!') {
                Some(p) => p,
                None => continue,
            };
            let range = part[bang + 1..].trim();
            if range.is_empty() {
                continue;
            }

            // Skip leading '$' characters when deciding whether this is a
            // row or column reference.
            let mut chars = range.chars().skip_while(|c| *c == '$');
            match chars.next() {
                Some(ch)
                    if ch.is_ascii_digit()
                        && rows.is_none()
                        && Self::is_valid_print_title_range(range, true) =>
                {
                    rows = Some(range.to_string());
                },
                Some(ch)
                    if ch.is_ascii_alphabetic()
                        && cols.is_none()
                        && Self::is_valid_print_title_range(range, false) =>
                {
                    cols = Some(range.to_string());
                },
                _ => {},
            }
        }

        (rows, cols)
    }

    fn split_defined_name_areas(value: &str) -> Vec<&str> {
        let bytes = value.as_bytes();
        let mut areas = Vec::new();
        let mut start = 0;
        let mut quoted = false;
        let mut index = 0;
        while index < bytes.len() {
            match bytes[index] {
                b'\'' if quoted && bytes.get(index + 1) == Some(&b'\'') => index += 1,
                b'\'' => quoted = !quoted,
                b',' if !quoted => {
                    areas.push(value[start..index].trim());
                    start = index + 1;
                },
                _ => {},
            }
            index += 1;
        }
        areas.push(value[start..].trim());
        areas
    }

    fn is_valid_print_cell_range(range: &str) -> bool {
        let mut references = range.split(':');
        let valid = references
            .by_ref()
            .take(2)
            .all(|reference| Cell::reference_to_coords(&reference.replace('$', "")).is_ok());
        valid && references.next().is_none()
    }

    fn is_valid_print_title_range(range: &str, rows: bool) -> bool {
        let mut endpoints = range.split(':');
        let mut count = 0;
        let valid = endpoints.by_ref().take(2).all(|endpoint| {
            count += 1;
            let endpoint = endpoint.trim_matches('$');
            let reference = if rows {
                format!("A{endpoint}")
            } else {
                format!("{endpoint}1")
            };
            Cell::reference_to_coords(&reference).is_ok()
        });
        valid && count == 2 && endpoints.next().is_none()
    }

    /// Get a worksheet by index
    fn get_worksheet(&self, index: usize) -> SheetResult<Worksheet<'_>> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        let info = &self.worksheets[index];
        let mut worksheet = Worksheet::new(self, info.clone());

        // Load worksheet data
        worksheet.load_data()?;

        Ok(worksheet)
    }

    /// Get the OPC package (for internal use by worksheet)
    pub(crate) fn package(&self) -> &OpcPackage {
        &self.package
    }

    /// Get the underlying OPC package.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.package
    }

    /// Get mutable OPC access, dropping signatures that would become stale.
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        let _ = self.package.clear_digital_signatures();
        &mut self.package
    }

    /// Verify package signatures without making a PKI trust determination.
    pub fn verify_digital_signatures(
        &self,
        policy: &litchi_opc::SignatureVerificationPolicy,
    ) -> litchi_opc::signature::Result<Vec<litchi_opc::DigitalSignatureVerification>> {
        self.package.verify_digital_signatures(policy)
    }

    /// Sign the current, fully materialized package while preserving valid signatures.
    pub fn add_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<PackURI> {
        self.package.add_digital_signature(signer)
    }

    /// Replace all package signatures with one new signature.
    pub fn resign_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<PackURI> {
        self.package.resign_digital_signature(signer)
    }

    /// Remove all package digital signatures.
    pub fn clear_digital_signatures(&mut self) -> litchi_opc::signature::Result<()> {
        self.package.clear_digital_signatures()
    }

    /// Resolve a worksheet through the relationship declared by workbook.xml.
    pub(crate) fn worksheet_part_uri(&self, worksheet: &WorksheetInfo) -> SheetResult<PackURI> {
        use litchi_opc::constants::relationship_type as rt;

        let workbook_part = self.package.get_part(&self.workbook_uri)?;
        let relationship = workbook_part
            .rels()
            .get(&worksheet.relationship_id)
            .ok_or_else(|| {
                format!(
                    "Worksheet '{}' references missing relationship '{}'",
                    worksheet.name, worksheet.relationship_id
                )
            })?;

        if relationship.reltype() != rt::WORKSHEET && relationship.reltype() != rt::STRICT_WORKSHEET
        {
            return Err(format!(
                "Relationship '{}' for worksheet '{}' has invalid type '{}'",
                worksheet.relationship_id,
                worksheet.name,
                relationship.reltype()
            )
            .into());
        }

        if relationship.is_external() {
            return Err(format!(
                "Relationship '{}' for worksheet '{}' has an external target",
                worksheet.relationship_id, worksheet.name
            )
            .into());
        }

        Ok(relationship.target_partname()?)
    }

    /// Whether the sheet at `index` is a worksheet rather than a
    /// chartsheet or dialogsheet, judged by its workbook relationship
    /// type (`sheet`, ECMA-376 §18.2.19; `chartsheet`, ECMA-376 §18.3.1.12).
    /// Sheet-scoped companion parts (ActiveX, web extensions) only exist
    /// on worksheets, so save-time detachment skips the rest.
    fn is_spreadsheetml_worksheet(&self, index: usize) -> bool {
        use litchi_opc::constants::relationship_type as rt;

        let (Ok(workbook_part), Some(info)) = (
            self.package.get_part(&self.workbook_uri),
            self.worksheets.get(index),
        ) else {
            return false;
        };
        workbook_part
            .rels()
            .get(&info.relationship_id)
            .is_some_and(|relationship| {
                relationship.reltype() == rt::WORKSHEET
                    || relationship.reltype() == rt::STRICT_WORKSHEET
            })
    }

    /// Get the shared strings table (for internal use by worksheet)
    pub(crate) fn shared_strings(&self) -> &SharedStrings {
        &self.shared_strings
    }

    /// Get the styles collection (for internal use by worksheet)
    pub(crate) fn styles(&self) -> &Styles {
        &self.styles
    }
}

impl WorkbookTrait for Workbook {
    fn active_worksheet(&self) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        let worksheet = self.get_worksheet(self.active_sheet_index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_names(&self) -> &[String] {
        // Return cached slice - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        for (index, ws_info) in self.worksheets.iter().enumerate() {
            if ws_info.name == name {
                let worksheet = self.get_worksheet(index)?;
                return Ok(Box::new(worksheet));
            }
        }
        Err(format!("Worksheet '{}' not found", name).into())
    }

    fn worksheet_by_index(&self, index: usize) -> SheetResult<Box<dyn WorksheetTrait + '_>> {
        let worksheet = self.get_worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(XlsxWorksheetIterator::new(self.worksheets.clone(), self))
    }

    fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        self.active_sheet_index
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904_date_system
    }
}

impl Workbook {
    /// Open a workbook from a path.
    pub fn open<P: AsRef<std::path::Path>>(path: P) -> SheetResult<Self> {
        let package = OpcPackage::open(path)?;
        Self::new(package)
    }

    #[cfg(feature = "encryption")]
    pub fn open_with_password<P: AsRef<std::path::Path>>(
        path: P,
        password: &str,
    ) -> SheetResult<Self> {
        let data = std::fs::read(path.as_ref())?;
        let decrypted = crate::crypto::decrypt_ooxml_if_encrypted(&data, password)?;
        let package = OpcPackage::from_bytes(&decrypted.package_bytes)?;
        Self::new(package)
    }

    /// Get a mutable worksheet for writing and modification.
    ///
    /// # Arguments
    ///
    /// * `index` - Worksheet index (0-based)
    ///
    // ... (rest of the code remains the same)
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// ws.set_cell_value(1, 1, "Hello");
    /// ws.set_cell_value(1, 2, "World");
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn worksheet_mut(&mut self, index: usize) -> SheetResult<&mut MutableWorksheet> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().worksheet_mut(index)
    }

    /// Return typed protection metadata, including any queued mutation.
    pub fn worksheet_protection_metadata(
        &self,
        index: usize,
    ) -> SheetResult<WorksheetProtectionMetadata> {
        if let Some(value) = self.worksheet_protection_mutations.get(&index) {
            return Ok(value.clone());
        }
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        parse_worksheet_protection(part.blob()).map_err(Into::into)
    }

    /// Atomically replace all worksheet protection metadata.
    pub fn replace_worksheet_protection(
        &mut self,
        index: usize,
        metadata: WorksheetProtectionMetadata,
    ) -> SheetResult<()> {
        validate_worksheet_protection_metadata(&metadata)?;
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        replace_worksheet_protection(part.blob(), &metadata)?;
        self.worksheet_protection_mutations.insert(index, metadata);
        Ok(())
    }

    /// Atomically update worksheet protection through a cloned candidate.
    pub fn update_worksheet_protection<F>(&mut self, index: usize, update: F) -> SheetResult<()>
    where
        F: FnOnce(&mut WorksheetProtectionMetadata),
    {
        let mut candidate = self.worksheet_protection_metadata(index)?;
        update(&mut candidate);
        self.replace_worksheet_protection(index, candidate)
    }

    pub fn set_sheet_protection(
        &mut self,
        index: usize,
        protection: WorksheetProtection,
    ) -> SheetResult<()> {
        let mut candidate = self.worksheet_protection_metadata(index)?;
        candidate.set_sheet_protection(Some(protection))?;
        self.replace_worksheet_protection(index, candidate)
    }

    pub fn remove_sheet_protection(&mut self, index: usize) -> SheetResult<()> {
        let mut candidate = self.worksheet_protection_metadata(index)?;
        candidate.clear_sheet_protection();
        self.replace_worksheet_protection(index, candidate)
    }

    pub fn set_protected_ranges(
        &mut self,
        index: usize,
        collections: Vec<WorksheetProtectedRangeCollection>,
    ) -> SheetResult<()> {
        let mut candidate = self.worksheet_protection_metadata(index)?;
        candidate.set_protected_range_collections(collections)?;
        self.replace_worksheet_protection(index, candidate)
    }

    pub fn remove_protected_ranges(&mut self, index: usize) -> SheetResult<()> {
        let mut candidate = self.worksheet_protection_metadata(index)?;
        candidate.clear_protected_ranges();
        self.replace_worksheet_protection(index, candidate)
    }

    /// Return complete core and Office 2010 data-validation collections.
    pub fn worksheet_data_validation_collections(
        &self,
        index: usize,
    ) -> SheetResult<Vec<DataValidationCollection>> {
        if let Some(value) = self.worksheet_data_validation_mutations.get(&index) {
            return Ok(value.clone());
        }
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        parse_data_validation_collections(part.blob()).map_err(Into::into)
    }

    /// Return all array formulas on one worksheet (`f` with `t="array"`,
    /// ECMA-376 §18.3.1.40), sorted in row-major order.
    pub fn worksheet_array_formulas(&self, index: usize) -> SheetResult<Vec<ArrayFormula>> {
        Ok(self.get_worksheet(index)?.get_array_formulas())
    }

    /// Atomically replace all data-validation collections on one existing worksheet.
    pub fn replace_worksheet_data_validations(
        &mut self,
        index: usize,
        collections: Vec<DataValidationCollection>,
    ) -> SheetResult<()> {
        validate_data_validation_collections(&collections)?;
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        replace_data_validation_collections(part.blob(), &collections)?;
        self.worksheet_data_validation_mutations
            .insert(index, collections);
        Ok(())
    }

    /// Atomically update cloned data-validation collections and commit only if valid.
    pub fn update_worksheet_data_validations<F>(
        &mut self,
        index: usize,
        update: F,
    ) -> SheetResult<()>
    where
        F: FnOnce(&mut Vec<DataValidationCollection>),
    {
        let mut candidate = self.worksheet_data_validation_collections(index)?;
        update(&mut candidate);
        self.replace_worksheet_data_validations(index, candidate)
    }

    /// Return inert Office Add-in range bindings, including any queued mutation.
    pub fn worksheet_web_extension_bindings(
        &self,
        index: usize,
    ) -> SheetResult<Vec<WorksheetWebExtensionBinding>> {
        if let Some(value) = self.worksheet_web_extension_binding_mutations.get(&index) {
            return Ok(value.clone());
        }
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        parse_worksheet_web_extension_bindings(part.blob())}

    /// Atomically replace all Office Add-in range bindings on one worksheet.
    ///
    /// Every non-empty worksheet `appRef` must resolve to exactly one binding
    /// in the package-level MS-OWEXML task-pane graph.
    pub fn replace_worksheet_web_extension_bindings(
        &mut self,
        index: usize,
        bindings: Vec<WorksheetWebExtensionBinding>,
    ) -> SheetResult<()> {
        let info = self
            .worksheets
            .get(index)
            .ok_or("Worksheet index out of bounds")?;
        let uri = self.worksheet_part_uri(info)?;
        let part = self.package.get_part(&uri)?;
        patch_worksheet_web_extension_bindings(part.blob(), &bindings)?;
        if !bindings.is_empty() {
            let task_panes = load_web_extension_task_panes(&self.package)?
                .ok_or("Worksheet add-in bindings require package task panes")?;
            let package_bindings = task_panes
                .panes
                .iter()
                .flat_map(|pane| pane.web_extension.bindings.iter().cloned())
                .collect::<Vec<_>>();
            validate_worksheet_web_extension_apprefs(&bindings, &package_bindings)?;
        }
        self.worksheet_web_extension_binding_mutations
            .insert(index, bindings);
        Ok(())
    }

    /// Atomically update cloned Office Add-in bindings and queue them if valid.
    pub fn update_worksheet_web_extension_bindings<F>(
        &mut self,
        index: usize,
        update: F,
    ) -> SheetResult<()>
    where
        F: FnOnce(&mut Vec<WorksheetWebExtensionBinding>),
    {
        let mut candidate = self.worksheet_web_extension_bindings(index)?;
        update(&mut candidate);
        self.replace_worksheet_web_extension_bindings(index, candidate)
    }

    /// Remove all worksheet-side Office Add-in bindings.
    pub fn remove_worksheet_web_extension_bindings(&mut self, index: usize) -> SheetResult<()> {
        self.replace_worksheet_web_extension_bindings(index, Vec::new())
    }

    pub fn remove_worksheet_data_validations(&mut self, index: usize) -> SheetResult<()> {
        self.replace_worksheet_data_validations(index, Vec::new())
    }

    /// Add a pivot table to the workbook (writer).
    ///
    /// This wires the pivot cache/table into the save pipeline; when you call
    /// `save`, the necessary parts and relationships will be created.
    pub fn add_pivot_table(&mut self, pivot: PivotTable) -> SheetResult<()> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().add_pivot_table(pivot)
    }

    /// Read the person list from the workbook.
    ///
    /// Persons are used in threaded comments to identify comment authors.
    /// Returns `None` if no person list is present in the workbook.
    pub fn persons(&self) -> SheetResult<Option<crate::xlsx::PersonList>> {
        crate::xlsx::read_persons(self.package())
    }

    /// Read threaded comments for a specific worksheet by index.
    ///
    /// Threaded comments are the modern comment format in Excel with support for
    /// conversation threads, @mentions, and timestamps.
    ///
    /// # Arguments
    /// * `sheet_index` - Zero-based worksheet index
    ///
    /// Returns `None` if the worksheet has no threaded comments.
    pub fn threaded_comments(
        &self,
        sheet_index: usize,
    ) -> SheetResult<Option<crate::xlsx::ThreadedComments>> {
        let ws_info = self
            .worksheets
            .get(sheet_index)
            .ok_or("Worksheet index out of bounds")?;

        let worksheet_uri = self.worksheet_part_uri(ws_info)?;
        let comments = crate::xlsx::read_threaded_comments(self.package(), &worksheet_uri)?;
        if let Some(comments) = comments.as_ref() {
            let people = self.persons()?;
            validate_threaded_comment_people(comments.comments.iter(), people.as_ref())?;
        }
        Ok(comments)
    }

    /// Read threaded comments for a worksheet by name.
    ///
    /// # Arguments
    /// * `sheet_name` - Name of the worksheet
    ///
    /// Returns `None` if the worksheet has no threaded comments.
    pub fn threaded_comments_by_name(
        &self,
        sheet_name: &str,
    ) -> SheetResult<Option<crate::xlsx::ThreadedComments>> {
        let sheet_index = self
            .worksheets
            .iter()
            .position(|ws| ws.name == sheet_name)
            .ok_or_else(|| format!("Worksheet '{}' not found", sheet_name))?;

        self.threaded_comments(sheet_index)
    }

    /// Add a new worksheet.
    ///
    /// # Arguments
    /// * `name` - The name of the new worksheet
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn add_worksheet(&mut self, name: &str) -> &mut MutableWorksheet {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .add_worksheet(name.to_string())
    }

    /// Insert a new worksheet at the requested workbook-order position
    /// (`sheet` inside `sheets`, ECMA-376 §18.2.19 and §18.2.20).
    ///
    /// `index` counts worksheets and chartsheets together in workbook
    /// order; passing the current sheet count appends like
    /// [`Self::add_worksheet`].
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.insert_worksheet(1, "Summary")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn insert_worksheet(
        &mut self,
        index: usize,
        name: &str,
    ) -> SheetResult<&mut MutableWorksheet> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .insert_worksheet(index, name.to_string())
    }

    /// Remove a worksheet by its worksheets-relative `index` and return it.
    ///
    /// The sheet leaves the `sheets` sequence (ECMA-376 §18.2.20); its
    /// part, workbook relationship, and content-type override are not
    /// emitted on the next save. Defined names scoped to the removed sheet
    /// (`definedName@localSheetId`, ECMA-376 §18.2.5) are dropped and
    /// later scopes shift up, and queued per-worksheet mutations
    /// (protection metadata, data validations, web-extension bindings)
    /// shift with it. Removing a sheet that a pivot table targets is
    /// rejected.
    ///
    /// Removal requires the writer data model: it fails for a workbook
    /// that was only opened for reading, because those sheets are not
    /// tracked as mutable worksheets.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Scratch");
    /// let removed = wb.remove_worksheet(1)?;
    /// assert_eq!(removed.name(), "Scratch");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn remove_worksheet(&mut self, index: usize) -> SheetResult<MutableWorksheet> {
        let data = self.mutable_data.as_mut().ok_or(
            "cannot remove a worksheet from a workbook opened read-only; \
             worksheet removal requires the writer data model",
        )?;
        let removed = data.remove_worksheet(index)?;
        shift_index_keyed_mutations(&mut self.worksheet_protection_mutations, index);
        shift_index_keyed_mutations(&mut self.worksheet_data_validation_mutations, index);
        shift_index_keyed_mutations(&mut self.worksheet_web_extension_binding_mutations, index);
        Ok(removed)
    }

    /// Add a chartsheet hosting the given chart.
    ///
    /// The chartsheet is appended to the workbook's sheet list in insertion
    /// order (interleaved with worksheets) and is emitted on save as a
    /// `/xl/chartsheets/` part with its own drawing and chart parts. The
    /// chart can be a classic chart or a pivot chart built with
    /// `WorksheetChart::into_pivot_chart`; the pivot-table binding is
    /// validated at save time like worksheet pivot charts.
    ///
    /// # Arguments
    /// * `name` - Chartsheet name (valid Excel sheet name, unique across sheets)
    /// * `chart` - The chart to host on the chartsheet
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, WorksheetChart, ChartAnchor};
    ///
    /// let mut wb = Workbook::create()?;
    /// let chart = WorksheetChart::bar_chart(
    ///     "Sales", "Sheet1!$A$2:$A$4", "Sheet1!$B$2:$B$4",
    ///     ChartAnchor::new(0, 0, 10, 15),
    /// )?;
    /// wb.add_chart_sheet("Sales Chart", chart)?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn add_chart_sheet(
        &mut self,
        name: &str,
        chart: crate::xlsx::WorksheetChart,
    ) -> SheetResult<&mut crate::xlsx::writer::MutableChartSheet> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .add_chart_sheet(name, chart)
    }

    /// Remove a chartsheet by its chartsheets-relative `index` and return it.
    ///
    /// Symmetric to [`Self::remove_worksheet`]: the chartsheet part, its
    /// drawing part, and the hosted chart part (`chartsheet`,
    /// ECMA-376 §18.3.1.12) are not emitted on the next save, and defined
    /// names scoped to its workbook position (`definedName@localSheetId`,
    /// ECMA-376 §18.2.5) are dropped with later scopes shifted up.
    ///
    /// Like [`Self::remove_worksheet`], removal requires the writer data
    /// model and fails for a workbook that was only opened for reading.
    pub fn remove_chart_sheet(
        &mut self,
        index: usize,
    ) -> SheetResult<crate::xlsx::writer::MutableChartSheet> {
        let data = self.mutable_data.as_mut().ok_or(
            "cannot remove a chartsheet from a workbook opened read-only; \
             chartsheet removal requires the writer data model",
        )?;
        data.remove_chart_sheet(index)
    }

    /// Add a new worksheet after validating its name.
    ///
    /// The name must satisfy Excel's sheet-name rules (1-31 characters,
    /// none of `: \ / ? * [ ]`) and be unique case-insensitively across
    /// worksheets and chartsheets — the same rules [`Self::add_chart_sheet`]
    /// enforces. [`Self::add_worksheet`] skips this validation for
    /// backward compatibility.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.try_add_worksheet("Summary")?;
    /// assert!(wb.try_add_worksheet("summary").is_err()); // duplicate
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn try_add_worksheet(&mut self, name: &str) -> SheetResult<&mut MutableWorksheet> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .try_add_worksheet(name.to_string())
    }

    /// Define a named range.
    ///
    /// Named ranges allow you to refer to cells or ranges by meaningful names.
    ///
    /// # Arguments
    /// * `name` - Name for the range (e.g., "TaxRate", "SalesData")
    /// * `reference` - Reference formula (e.g., "Sheet1!$A$1:$B$10", "Sheet1!$C$5")
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.define_name("TaxRate", "Sheet1!$A$1");
    /// wb.define_name("SalesData", "Sheet1!$A$1:$D$100");
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn define_name(&mut self, name: &str, reference: &str) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name(name, reference);
    }

    /// Define a sheet-scoped named range.
    ///
    /// # Arguments
    /// * `name` - Name for the range
    /// * `reference` - Reference formula
    /// * `sheet_id` - Zero-based workbook sheet index (`localSheetId` in OOXML)
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet_id: u32) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name_local(name, reference, sheet_id);
    }

    /// Define a named range with a comment.
    pub fn define_name_with_comment(&mut self, name: &str, reference: &str, comment: &str) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .define_name_with_comment(name, reference, comment);
    }

    /// Remove a named range by name.
    pub fn remove_name(&mut self, name: &str) -> bool {
        self.mutable_data
            .as_mut()
            .map(|d| d.remove_name(name))
            .unwrap_or(false)
    }

    /// Get a reference to the workbook properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::create()?;
    /// let props = wb.properties();
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn properties(&self) -> &DocumentProperties {
        &self.properties
    }

    /// Get a mutable reference to the workbook properties.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.properties_mut().title = Some("My Workbook".to_string());
    /// wb.properties_mut().creator = Some("John Doe".to_string());
    /// wb.save("workbook.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn properties_mut(&mut self) -> &mut DocumentProperties {
        &mut self.properties
    }

    /// Set the person list for threaded comments.
    ///
    /// Persons are used to identify authors of threaded comments.
    ///
    /// # Arguments
    /// * `person_list` - List of persons who can author comments
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, Person, PersonList};
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut person_list = PersonList::default();
    /// person_list.persons.push(Person {
    ///     display_name: "John Doe".to_string(),
    ///     id: "{11111111-2222-3333-4444-555555555555}".to_string(),
    ///     user_id: None,
    ///     provider_id: None,
    /// });
    /// wb.set_person_list(person_list);
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_person_list(&mut self, person_list: crate::xlsx::PersonList) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        if let Some(ref mut data) = self.mutable_data {
            data.person_list = Some(person_list);
        }
    }

    /// Save the workbook to a file.
    ///
    /// Writes the complete Excel workbook including all worksheets, styles,
    /// and shared strings to an .xlsx file.
    ///
    /// # Arguments
    /// * `path` - Path where the .xlsx file should be written
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut workbook = Workbook::create()?;
    /// // Modify workbook...
    /// workbook.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> SheetResult<()> {
        // If we have mutable data, update the workbook parts
        let should_update = self
            .mutable_data
            .as_ref()
            .map(|d| d.is_modified())
            .unwrap_or(false);

        if should_update {
            // Take mutable_data temporarily to avoid borrow issues
            if let Some(mut mutable_data) = self.mutable_data.take() {
                let worksheet_web_extension_bindings = (0..self.worksheets.len())
                    .filter(|index| self.is_spreadsheetml_worksheet(*index))
                    .map(|index| {
                        let bindings = self.worksheet_web_extension_bindings(index)?;
                        Ok((!bindings.is_empty()
                            || self
                                .worksheet_web_extension_binding_mutations
                                .contains_key(&index))
                        .then_some((index, bindings)))
                    })
                    .collect::<SheetResult<Vec<_>>>()?
                    .into_iter()
                    .flatten()
                    .collect::<Vec<_>>();
                // The mutable writer rebuilds workbook and worksheet relationship
                // collections. Detach inert companion parts first so old targets
                // do not become orphaned, then restore them after materialization.
                let named_sheet_views = self.detach_named_sheet_views_before_materialization()?;
                let active_x_controls = (0..self.worksheets.len())
                    .filter(|index| self.is_spreadsheetml_worksheet(*index))
                    .map(|index| {
                        let info = self
                            .worksheets
                            .get(index)
                            .ok_or("Worksheet index out of bounds")?;
                        let uri = self.worksheet_part_uri(info)?;
                        let controls = load_worksheet_active_x(&self.package, &uri)?;
                        Ok((!controls.controls.is_empty()).then_some((uri, controls)))
                    })
                    .collect::<SheetResult<Vec<_>>>()?
                    .into_iter()
                    .flatten()
                    .collect::<Vec<_>>();
                for (uri, _) in &active_x_controls {
                    remove_worksheet_active_x(&mut self.package, uri)?;
                }
                let volatile_dependencies = load_volatile_dependencies(&self.package)?;
                let xml_maps = load_xml_maps(&self.package)?;
                if self.calculation_chain.is_some() {
                    remove_calculation_chain(&mut self.package)?;
                }
                if volatile_dependencies.is_some() {
                    remove_volatile_dependencies(&mut self.package)?;
                }
                if xml_maps.is_some() {
                    remove_xml_maps(&mut self.package)?;
                }
                self.update_workbook_parts(&mut mutable_data)?;
                self.restore_calculation_chain_after_materialization()?;
                self.restore_volatile_dependencies_after_materialization(&volatile_dependencies)?;
                self.restore_xml_maps_after_materialization(&xml_maps)?;
                self.restore_named_sheet_views_after_materialization(&named_sheet_views)?;
                for (worksheet_index, (uri, mut controls)) in
                    active_x_controls.into_iter().enumerate()
                {
                    let mut occupied = self
                        .package
                        .get_part(&uri)?
                        .rels()
                        .iter()
                        .map(|relationship| relationship.r_id().to_string())
                        .collect::<HashSet<_>>();
                    for (control_index, item) in controls.controls.iter_mut().enumerate() {
                        item.descriptor_uri = next_active_x_part_uri(
                            &self.package,
                            "/xl/activeX",
                            &format!("litchiControl{worksheet_index}_{control_index}"),
                            "xml",
                        )?;
                        for (binary_index, binary) in item.binaries.iter_mut().enumerate() {
                            binary.part_uri = next_active_x_part_uri(
                                &self.package,
                                "/xl/activeX",
                                &format!(
                                    "litchiControl{worksheet_index}_{control_index}Binary{binary_index}"
                                ),
                                "bin",
                            )?;
                        }
                        item.control.relationship_id =
                            next_active_x_relationship_id(&mut occupied, control_index, false);
                        if let Some(preview) = item.preview.as_mut() {
                            preview.part_uri = next_active_x_part_uri(
                                &self.package,
                                "/xl/media",
                                &format!(
                                    "litchiControl{worksheet_index}_{control_index}Preview"
                                ),
                                "img",
                            )?;
                            let id =
                                next_active_x_relationship_id(&mut occupied, control_index, true);
                            preview.relationship_id = id.clone();
                            if let Some(properties) = item.control.properties.as_mut() {
                                properties.preview_relationship_id = Some(id);
                            }
                        }
                    }
                    store_worksheet_active_x(&mut self.package, &uri, &controls)?;
                }
                self.restore_worksheet_web_extension_bindings_after_materialization(
                    &mutable_data,
                    &worksheet_web_extension_bindings,
                )?;
                self.mutable_data = Some(mutable_data);

                // Re-sync the read-side model (worksheet relationship IDs,
                // defined names, shared strings, styles) with the parts
                // just written, so a second save resolves the same state
                // as a reopened workbook.
                self.load_workbook_info()?;
                self.load_shared_strings()?;
                self.load_styles()?;
            }
        }

        // Update core properties
        self.update_core_properties()?;

        // Update app properties (extended properties)
        self.update_app_properties()?;

        let staged = self.stage_worksheet_mutations()?;
        for (uri, _, replacement) in &staged {
            self.package
                .get_part_mut(uri)?
                .set_blob(replacement.clone());
        }
        let save_result = self.package.save(path);
        for (uri, original, _) in staged {
            self.package
                .get_part_mut(&uri)
                .expect("staged worksheet part remains present")
                .set_blob(original);
        }
        save_result?;
        Ok(())
    }

    #[allow(clippy::type_complexity)]
    fn stage_worksheet_mutations(&self) -> SheetResult<Vec<(PackURI, Vec<u8>, Vec<u8>)>> {
        use litchi_opc::constants::content_type as ct;
        let indexes: HashSet<usize> = self
            .worksheet_protection_mutations
            .keys()
            .chain(self.worksheet_data_validation_mutations.keys())
            .chain(self.worksheet_web_extension_binding_mutations.keys())
            .copied()
            .collect();
        let mut staged = Vec::with_capacity(indexes.len());
        for index in indexes {
            let info = self
                .worksheets
                .get(index)
                .ok_or("Worksheet index out of bounds")?;
            let uri = self.worksheet_part_uri(info)?;
            let part = self.package.get_part(&uri)?;
            if part.content_type() != ct::SML_WORKSHEET {
                return Err(format!(
                    "Worksheet '{}' has invalid content type '{}'",
                    info.name,
                    part.content_type()
                )
                .into());
            }
            let original = part.blob().to_vec();
            let mut replacement = original.clone();
            if let Some(metadata) = self.worksheet_protection_mutations.get(&index) {
                replacement = replace_worksheet_protection(&replacement, metadata)?;
            }
            if let Some(collections) = self.worksheet_data_validation_mutations.get(&index) {
                replacement = replace_data_validation_collections(&replacement, collections)?;
            }
            if let Some(bindings) = self.worksheet_web_extension_binding_mutations.get(&index) {
                replacement = patch_worksheet_web_extension_bindings(&replacement, bindings)?;
            }
            staged.push((uri, original, replacement));
        }
        Ok(staged)
    }

    /// The writer rebuilds the workbook relationships; restore the optional
    /// inert cache afterwards without recalculating formulas.
    fn restore_calculation_chain_after_materialization(&mut self) -> SheetResult<()> {
        let Some(chain) = self.calculation_chain.as_ref() else {
            return Ok(());
        };
        store_calculation_chain(
            &mut self.package,
            chain,
            self.calculation_chain_conformance.unwrap_or_default(),
        )?;
        Ok(())
    }

    fn restore_worksheet_web_extension_bindings_after_materialization(
        &mut self,
        data: &MutableWorkbookData,
        bindings_by_index: &[(usize, Vec<WorksheetWebExtensionBinding>)],
    ) -> SheetResult<()> {
        for (index, bindings) in bindings_by_index {
            let worksheet = data
                .worksheets
                .get(*index)
                .ok_or("Worksheet index out of bounds after materialization")?;
            let uri = PackURI::new(format!("/xl/worksheets/sheet{}.xml", worksheet.sheet_id()))?;
            let part = self.package.get_part_mut(&uri)?;
            let replacement = patch_worksheet_web_extension_bindings(part.blob(), bindings)?;
            part.set_blob(replacement);
            self.worksheet_web_extension_binding_mutations.remove(index);
        }
        Ok(())
    }

    /// Restore Custom XML Maps after the mutable writer recreates the workbook
    /// relationship collection. No mapping is applied during this operation.
    fn restore_xml_maps_after_materialization(
        &mut self,
        value: &Option<(XmlMapInfo, XmlMapConformance)>,
    ) -> SheetResult<()> {
        let Some((value, conformance)) = value else {
            return Ok(());
        };
        store_xml_maps(&mut self.package, value, *conformance)
    }

    /// Restore volatile-dependencies metadata after the mutable writer rebuilds
    /// the workbook relationship collection without refreshing any dependency.
    fn restore_volatile_dependencies_after_materialization(
        &mut self,
        value: &Option<(VolatileDependencies, VolatileDependenciesConformance)>,
    ) -> SheetResult<()> {
        let Some((value, conformance)) = value else {
            return Ok(());
        };
        store_volatile_dependencies(&mut self.package, value, *conformance)
    }

    /// Detach worksheet-scoped modern views before rebuilding worksheet parts.
    /// The mutable writer owns its relationship collections, so leaving these
    /// parts attached would create orphaned package data.
    fn detach_named_sheet_views_before_materialization(
        &mut self,
    ) -> SheetResult<Vec<(u32, NamedSheetViews)>> {
        // Named sheet views only exist on worksheets; chartsheets share the
        // sheets sequence but have no worksheet part to detach from.
        let mut worksheets = Vec::new();
        for (index, info) in self.worksheets.iter().enumerate() {
            if self.is_spreadsheetml_worksheet(index) {
                worksheets.push((info.sheet_id, self.worksheet_part_uri(info)?));
            }
        }
        let mut retained = Vec::new();
        for (sheet_id, worksheet_part) in worksheets {
            if let Some(value) = load_worksheet_named_sheet_views(&self.package, &worksheet_part)? {
                remove_worksheet_named_sheet_views(&mut self.package, &worksheet_part)?;
                retained.push((sheet_id, value));
            }
        }
        Ok(retained)
    }

    /// Restore parsed Named Sheet Views after the mutable writer has recreated
    /// worksheet parts and their relationship collections.
    fn restore_named_sheet_views_after_materialization(
        &mut self,
        retained: &[(u32, NamedSheetViews)],
    ) -> SheetResult<()> {
        for (sheet_id, value) in retained {
            let worksheet_part = PackURI::new(format!("/xl/worksheets/sheet{sheet_id}.xml"))?;
            if self.package.get_part(&worksheet_part).is_ok() {
                store_worksheet_named_sheet_views(&mut self.package, &worksheet_part, value)?;
            }
        }
        Ok(())
    }

    /// Generate metadata.xml for threaded comments support.
    fn generate_metadata_xml() -> String {
        xml_minifier::minified_xml!("resources/metadata.xml").to_string()
    }

    /// Generate bridge comments.xml for threaded comments backwards compatibility.
    fn generate_bridge_comments_xml(threaded_comments: &crate::xlsx::ThreadedComments) -> String {
        use litchi_core::xml::escape::escape_xml;
        use std::collections::HashMap;

        let mut xml = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xml.push_str(
            "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
        );

        // Build author list and comment grouping
        let mut authors = Vec::new();
        let mut author_map: HashMap<String, usize> = HashMap::new();
        let mut comments_by_cell: HashMap<String, Vec<&crate::xlsx::ThreadedComment>> =
            HashMap::new();

        // Group comments by cell
        for comment in &threaded_comments.comments {
            if let Some(ref cell_ref) = comment.cell_ref {
                comments_by_cell
                    .entry(cell_ref.clone())
                    .or_default()
                    .push(comment);
            }
        }

        // Build authors list with tc={id} entries
        for cell_comments in comments_by_cell.values() {
            for comment in cell_comments {
                if !author_map.contains_key(&comment.id) {
                    let author_entry = format!("tc={}", comment.id);
                    author_map.insert(comment.id.clone(), authors.len());
                    authors.push(author_entry);
                }
            }
        }

        // Write authors
        xml.push_str("<authors>");
        for author in &authors {
            xml.push_str(&format!("<author>{}</author>", escape_xml(author)));
        }
        xml.push_str("</authors>");

        // Write comment list
        xml.push_str("<commentList>");
        for (cell_ref, cell_comments) in comments_by_cell {
            if let Some(first_comment) = cell_comments.first() {
                let author_id = author_map.get(&first_comment.id).unwrap_or(&0);

                // Build combined text from all comments in thread
                let mut combined_text = String::new();
                for (i, comment) in cell_comments.iter().enumerate() {
                    if i > 0 {
                        combined_text.push_str("\nReply:\n    ");
                    } else {
                        combined_text.push_str("Comment:\n    ");
                    }
                    if let Some(ref text) = comment.text {
                        combined_text.push_str(text);
                    }
                }

                xml.push_str(&format!(
                    "<comment ref=\"{}\" authorId=\"{}\"><text><r><t>{}</t></r></text></comment>",
                    escape_xml(&cell_ref),
                    author_id,
                    escape_xml(&combined_text)
                ));
            }
        }
        xml.push_str("</commentList>");
        xml.push_str("</comments>");
        xml
    }

    /// Update workbook parts with modified data.
    /// Validate an authored pivot chart's binding against the workbook's
    /// pivot tables, returning a normalized copy with the canonical
    /// sheet-qualified pivot-source name. Returns `None` for ordinary charts.
    fn normalized_pivot_chart_model(
        chart: &crate::charts::Chart,
        host_sheet_name: &str,
        authored_pivot_tables: &[(String, String)],
    ) -> SheetResult<Option<crate::charts::Chart>> {
        if chart.pivot_source.is_none() {
            return Ok(None);
        }
        let mut normalized = chart.clone();
        let pivot_source = normalized
            .pivot_source
            .as_mut()
            .expect("pivot source presence checked above");
        pivot_source.name = crate::xlsx::pivot_chart::resolve_authored_pivot_source_name(
            &pivot_source.name,
            host_sheet_name,
            authored_pivot_tables,
        )?;
        Ok(Some(normalized))
    }

    /// Remove writer-owned sheet parts left over from a previous
    /// materialization (for example a sheet removed between two saves),
    /// so saving twice never emits orphaned parts. Live parts are derived
    /// from the writer data (`sheet`, ECMA-376 §18.2.19; `chartsheet`,
    /// ECMA-376 §18.3.1.12); anything under the writer-owned naming
    /// schemes without live backing is dropped.
    fn remove_stale_sheet_parts(&mut self, data: &MutableWorkbookData) {
        let worksheet_ids: HashSet<u32> =
            data.worksheets.iter().map(|ws| ws.sheet_id()).collect();
        let all_sheet_ids: HashSet<u32> = worksheet_ids
            .iter()
            .copied()
            .chain(data.chart_sheets.iter().map(|sheet| sheet.sheet_id()))
            .collect();
        let chartsheet_count = data.chart_sheets.len();

        let stale: Vec<PackURI> = self
            .package
            .iter_parts()
            .map(|part| part.partname().clone())
            .filter(|uri| {
                is_stale_sheet_part(uri.as_str(), &worksheet_ids, &all_sheet_ids, chartsheet_count)
            })
            .collect();
        for uri in stale {
            self.package.remove_part(&uri);
        }
    }

    fn update_workbook_parts(&mut self, data: &mut MutableWorkbookData) -> SheetResult<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::constants::relationship_type as rt;
        use litchi_opc::part::{BlobPart, Part};

        validate_workbook_tables(data)?;
        self.remove_stale_sheet_parts(data);

        let (
            preserved_main_content_type,
            preserved_vba_target,
            preserved_external_relationships,
        ) = {
            let workbook_part = self.package.get_part(&self.workbook_uri)?;
            discover_vba_project(&self.package, workbook_part)?;
            let mut vba_projects = workbook_part
                .rels()
                .iter()
                .filter(|relationship| relationship.reltype() == rt::VBA_PROJECT);
            let preserved_vba_target = match vba_projects.next() {
                Some(relationship) if relationship.is_external() => {
                    return Err("workbook VBA Project relationship cannot be external".into());
                },
                Some(relationship) => Some(relationship.target_ref().to_string()),
                None => None,
            };
            if vba_projects.next().is_some() {
                return Err("workbook has multiple VBA Project relationships".into());
            }
            let external_relationships = self.external_links
                .iter()
                .map(|link| {
                    let relationship =
                        workbook_part
                            .rels()
                            .get(&link.relationship_id)
                            .ok_or_else(|| {
                                format!(
                                    "missing preserved external-link relationship '{}'",
                                    link.relationship_id
                                )
                            })?;
                    Ok((
                        relationship.reltype().to_string(),
                        relationship.target_ref().to_string(),
                        relationship.r_id().to_string(),
                        relationship.is_external(),
                    ))
                })
                .collect::<SheetResult<Vec<_>>>()?;
            (
                workbook_part.content_type().to_string(),
                preserved_vba_target,
                external_relationships,
            )
        };

        let workbook_uri = PackURI::new("/xl/workbook.xml")?;

        // Create temporary workbook part to manage relationships
        let mut temp_wb_part = BlobPart::new(
            workbook_uri.clone(),
            preserved_main_content_type,
            Vec::new(),
        );
        for (relationship_type, target, relationship_id, external) in
            preserved_external_relationships
        {
            temp_wb_part.rels_mut().add_relationship(
                relationship_type,
                target,
                relationship_id,
                external,
            );
        }
        if let Some(target) = preserved_vba_target {
            temp_wb_part.relate_to(&target, rt::VBA_PROJECT);
        }

        // Build styles from all worksheets FIRST
        let (styles_builder, worksheet_style_indices) = data.build_styles()?;

        // Generate and write styles.xml
        let styles_xml = styles_builder.to_xml()?;
        let styles_uri = PackURI::new("/xl/styles.xml")?;
        let styles_part = BlobPart::new(
            styles_uri,
            ct::SML_STYLES.to_string(),
            styles_xml.into_bytes(),
        );
        self.package.add_part(Box::new(styles_part));

        // Create styles relationship
        temp_wb_part.relate_to("styles.xml", rt::STYLES);

        // Create theme relationship (required by Excel)
        temp_wb_part.relate_to("theme/theme1.xml", rt::THEME);

        // Track worksheet relationship IDs for workbook.xml generation
        let mut worksheet_rel_ids: Vec<String> = Vec::new();

        // Track pivot cache relationship IDs for workbook.xml
        let mut pivot_cache_rel_ids: Vec<(u32, String)> = Vec::new();
        // Track pivot table targets per worksheet (for worksheet rels)
        let mut pivot_table_targets_per_sheet: Vec<Vec<String>> =
            vec![Vec::new(); data.worksheets.len()];

        // Pre-create pivot cache and pivot table parts so worksheets can relate to them
        for (idx, pivot) in data.pivot_tables.iter().enumerate() {
            let cache_id = (idx as u32) + 1;

            // pivotCacheRecords part (materialized from source range)
            let records_uri =
                PackURI::new(format!("/xl/pivotCache/pivotCacheRecords{}.xml", cache_id))?;
            let (records_xml, record_count, field_stats) =
                generate_pivot_cache_records_xml(pivot, &data.worksheets)?;
            let records_part = BlobPart::new(
                records_uri,
                ct::SML_PIVOT_CACHE_RECORDS.to_string(),
                records_xml.into_bytes(),
            );
            self.package.add_part(Box::new(records_part));

            // pivotCacheDefinition part
            let cache_def_uri = PackURI::new(format!(
                "/xl/pivotCache/pivotCacheDefinition{}.xml",
                cache_id
            ))?;
            let mut cache_def_part = BlobPart::new(
                cache_def_uri,
                ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
                Vec::new(),
            );
            let records_rel_id = cache_def_part.relate_to(
                &format!("pivotCacheRecords{}.xml", cache_id),
                rt::PIVOT_CACHE_RECORDS,
            );
            let cache_def_xml = generate_pivot_cache_definition_xml(
                pivot,
                Some(records_rel_id.as_str()),
                record_count,
                &field_stats,
            )?;
            cache_def_part.set_blob(cache_def_xml.into_bytes());
            self.package.add_part(Box::new(cache_def_part));

            // workbook -> pivotCacheDefinition rel
            let cache_rel_id = temp_wb_part.relate_to(
                &format!("pivotCache/pivotCacheDefinition{}.xml", cache_id),
                rt::PIVOT_CACHE_DEFINITION,
            );
            pivot_cache_rel_ids.push((cache_id, cache_rel_id.clone()));

            // pivotTableDefinition part
            let table_idx = cache_id; // align ids for predictability
            let pivot_table_uri =
                PackURI::new(format!("/xl/pivotTables/pivotTable{}.xml", table_idx))?;
            let mut pivot_table_part =
                BlobPart::new(pivot_table_uri, ct::SML_PIVOT_TABLE.to_string(), Vec::new());

            // pivotTable -> pivotCacheDefinition rel
            let _pt_cache_rel_id = pivot_table_part.relate_to(
                &format!("../pivotCache/pivotCacheDefinition{}.xml", cache_id),
                rt::PIVOT_CACHE_DEFINITION,
            );

            // Serialize pivotTable XML
            let pivot_table_xml =
                generate_pivot_table_definition_xml(pivot, cache_id, &field_stats)?;
            pivot_table_part.set_blob(pivot_table_xml.into_bytes());
            self.package.add_part(Box::new(pivot_table_part));

            // Record worksheet target for later worksheet rel creation
            let sheet_idx = pivot.dest_sheet_index;
            if let Some(list) = pivot_table_targets_per_sheet.get_mut(sheet_idx) {
                list.push(format!("../pivotTables/pivotTable{}.xml", table_idx));
            } else {
                return Err(format!(
                    "Pivot table destination sheet index {} out of bounds",
                    sheet_idx
                )
                .into());
            }
        }

        // Materialize the pivot output into destination worksheet cells.
        // This ensures Excel shows the pivot table content immediately on open.
        for pivot in data.pivot_tables.iter() {
            render_pivot_table_sheet_cells(pivot, &mut data.worksheets)?;
        }

        // Write person list and metadata if present (for threaded comments)
        let has_threaded_comments = data
            .worksheets
            .iter()
            .any(|ws| !ws.threaded_comments().is_empty());
        validate_threaded_comment_people(
            data.worksheets
                .iter()
                .flat_map(|worksheet| worksheet.threaded_comments()),
            data.person_list.as_ref(),
        )?;

        if let Some(person_list) = data.person_list.as_ref()
            && !person_list.persons.is_empty()
        {
            let persons_xml = crate::xlsx::write_persons(person_list)?;
            let persons_uri = PackURI::new("/xl/persons/person.xml")?;
            let persons_part = BlobPart::new(
                persons_uri,
                ct::SML_PERSONS.to_string(),
                persons_xml.into_bytes(),
            );
            self.package.add_part(Box::new(persons_part));

            // Add relationship from workbook to persons part
            temp_wb_part.relate_to("persons/person.xml", rt::PERSONS);
        }

        // Write metadata.xml if there are threaded comments
        if has_threaded_comments {
            let metadata_xml = Self::generate_metadata_xml();
            // ...
            let metadata_uri = PackURI::new("/xl/metadata.xml")?;
            let metadata_part = BlobPart::new(
                metadata_uri,
                ct::SML_SHEET_METADATA.to_string(),
                metadata_xml.into_bytes(),
            );
            self.package.add_part(Box::new(metadata_part));

            // Create relationship from workbook to metadata
            temp_wb_part.relate_to("metadata.xml", rt::SHEET_METADATA);
        }

        // Authored pivot tables as (name, hosting sheet name) pairs, used to
        // validate authored pivot-chart bindings in the chart loop below.
        let authored_pivot_tables: Vec<(String, String)> = data
            .pivot_tables
            .iter()
            .map(|pivot| {
                (
                    pivot.name.clone(),
                    data.worksheets
                        .get(pivot.dest_sheet_index)
                        .map(|worksheet| worksheet.name().to_string())
                        .unwrap_or_default(),
                )
            })
            .collect();

        // Update worksheet parts and create relationships
        // IMPORTANT: Create relationships for ALL worksheets, not just modified ones
        for (index, ws) in data.worksheets.iter().enumerate() {
            // Get style indices for this worksheet
            let style_indices = worksheet_style_indices
                .get(index)
                .cloned()
                .unwrap_or_default();

            let ws_uri = PackURI::new(format!("/xl/worksheets/sheet{}.xml", ws.sheet_id()))?;

            // Create worksheet part with empty content initially (we'll set it later)
            let mut ws_part =
                BlobPart::new(ws_uri.clone(), ct::SML_WORKSHEET.to_string(), Vec::new());

            // Generate and add comments if present, create relationship
            if let Some(comments_xml) = ws.generate_comments_xml()? {
                let comments_uri = PackURI::new(format!("/xl/comments{}.xml", ws.sheet_id()))?;
                let comments_part = BlobPart::new(
                    comments_uri,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
                        .to_string(),
                    comments_xml.into_bytes(),
                );
                self.package.add_part(Box::new(comments_part));

                // Add relationship from worksheet to comments
                ws_part.relate_to(
                    &format!("../comments{}.xml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                );
            }

            // Generate and add threaded comments if present
            if !ws.threaded_comments().is_empty() {
                let threaded_comments = crate::xlsx::ThreadedComments {
                    comments: ws.threaded_comments().to_vec(),
                };
                let tc_xml = crate::xlsx::write_threaded_comments(&threaded_comments)?;
                let tc_uri = PackURI::new(format!(
                    "/xl/threadedComments/threadedComment{}.xml",
                    ws.sheet_id()
                ))?;
                let tc_part = BlobPart::new(
                    tc_uri,
                    ct::SML_THREADED_COMMENTS.to_string(),
                    tc_xml.into_bytes(),
                );
                self.package.add_part(Box::new(tc_part));

                // Add relationship from worksheet to threaded comments
                ws_part.relate_to(
                    &format!("../threadedComments/threadedComment{}.xml", ws.sheet_id()),
                    rt::THREADED_COMMENTS,
                );

                // Generate bridge comments.xml for backwards compatibility
                let bridge_xml = Self::generate_bridge_comments_xml(&threaded_comments);
                let bridge_uri = PackURI::new(format!("/xl/comments{}.xml", ws.sheet_id()))?;
                let bridge_part = BlobPart::new(
                    bridge_uri,
                    ct::SML_COMMENTS.to_string(),
                    bridge_xml.into_bytes(),
                );
                self.package.add_part(Box::new(bridge_part));

                // Add relationship from worksheet to bridge comments
                ws_part.relate_to(
                    &format!("../comments{}.xml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                );
            }

            // Generate and add VML drawing for comment indicators if present
            let vml_rel_id = if let Some(vml_xml) = ws.generate_vml_drawing_xml()? {
                let vml_uri =
                    PackURI::new(format!("/xl/drawings/vmlDrawing{}.vml", ws.sheet_id()))?;
                let vml_part = BlobPart::new(
                    vml_uri,
                    "application/vnd.openxmlformats-officedocument.vmlDrawing".to_string(),
                    vml_xml.into_bytes(),
                );
                self.package.add_part(Box::new(vml_part));

                // Add relationship from worksheet to VML drawing and capture the ID
                let rel_id = ws_part.relate_to(
                    &format!("../drawings/vmlDrawing{}.vml", ws.sheet_id()),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                );
                Some(rel_id)
            } else {
                None
            };

            // Add relationships for external hyperlinks and track their IDs
            let mut hyperlink_rel_ids: HashMap<String, String> = HashMap::new();
            for hyperlink in ws.hyperlinks().iter() {
                if hyperlink.target.starts_with("http://")
                    || hyperlink.target.starts_with("https://")
                    || hyperlink.target.starts_with("ftp://")
                    || hyperlink.target.starts_with("mailto:")
                {
                    // Use relate_to_ext for external links to add TargetMode="External"
                    let rel_id = ws_part.relate_to_ext(
                        &hyperlink.target,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    );
                    hyperlink_rel_ids.insert(hyperlink.cell_ref.clone(), rel_id);
                }
            }

            // Generate and add table XML files if present and track relationship IDs
            let mut table_rel_ids: Vec<String> = Vec::new();
            for table in ws.tables().iter() {
                use crate::xlsx::writer::table::serialize_table;

                let table_xml = serialize_table(table)?;
                let table_uri = PackURI::new(format!("/xl/tables/table{}.xml", table.id))?;
                let table_part =
                    BlobPart::new(table_uri, ct::SML_TABLE.to_string(), table_xml.into_bytes());
                self.package.add_part(Box::new(table_part));

                // Add relationship from worksheet to table and capture the ID
                let rel_id =
                    ws_part.relate_to(&format!("../tables/table{}.xml", table.id), rt::TABLE);
                table_rel_ids.push(rel_id);
            }

            // Generate and add drawing XML for images if present
            let drawing_rel_id = if let Some(drawing_xml) = ws.generate_drawing_xml()? {
                let drawing_uri =
                    PackURI::new(format!("/xl/drawings/drawing{}.xml", ws.sheet_id()))?;

                // Create drawing part with relationships for images
                let mut drawing_part = BlobPart::new(
                    drawing_uri.clone(),
                    ct::OFC_DRAWING.to_string(),
                    drawing_xml.into_bytes(),
                );

                // Add image parts and create relationships
                for (idx, image) in ws.images().iter().enumerate() {
                    let image_ext = &image.format;
                    let image_uri = PackURI::new(format!(
                        "/xl/media/image{}_{}.{}",
                        ws.sheet_id(),
                        idx + 1,
                        image_ext
                    ))?;

                    // Determine content type based on format
                    let content_type = match image_ext.to_lowercase().as_str() {
                        "png" => "image/png",
                        "jpg" | "jpeg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        "svg" => "image/svg+xml",
                        _ => {
                            return Err(format!(
                                "Unsupported worksheet image format '{}'",
                                image.format
                            )
                            .into());
                        },
                    };

                    let image_part = BlobPart::new(
                        image_uri.clone(),
                        content_type.to_string(),
                        image.data.clone(),
                    );
                    self.package.add_part(Box::new(image_part));

                    // Add relationship from drawing to image
                    drawing_part.relate_to(
                        &format!("../media/image{}_{}.{}", ws.sheet_id(), idx + 1, image_ext),
                        rt::IMAGE,
                    );
                }

                // Add chart parts and create relationships
                for (idx, chart) in ws.charts().iter().enumerate() {
                    let chart_name = format!("chart{}_{}", ws.sheet_id(), idx + 1);
                    let chart_uri = PackURI::new(format!("/xl/charts/{chart_name}.xml"))?;
                    if chart.chart.external_data.is_some() != chart.external_data_part.is_some() {
                        return Err(format!(
                            "Worksheet chart {} external-data metadata and package payload disagree",
                            idx + 1
                        )
                        .into());
                    }
                    if chart.chart.user_shapes.is_some() != chart.user_shapes_part.is_some() {
                        return Err(format!(
                            "Worksheet chart {} user-shapes metadata and package payload disagree",
                            idx + 1
                        )
                        .into());
                    }

                    // Validate authored pivot-chart bindings against the
                    // workbook's pivot tables and normalize the pivot-source
                    // name to its sheet-qualified form, so saved packages
                    // are valid by construction.
                    let normalized_chart = Self::normalized_pivot_chart_model(
                        &chart.chart,
                        ws.name(),
                        &authored_pivot_tables,
                    )?;
                    let chart_model = normalized_chart.as_ref().unwrap_or(&chart.chart);

                    let mut chart_part =
                        BlobPart::new(chart_uri.clone(), ct::DML_CHART.to_string(), Vec::new());
                    let mut chart_related_resources = Vec::new();
                    let mut additional_relationship_ids = std::collections::HashSet::new();
                    for (relationship_index, relationship) in
                        chart.additional_relationships.iter().enumerate()
                    {
                        if relationship.relationship_id.is_empty()
                            || relationship.relationship_type.is_empty()
                            || !additional_relationship_ids
                                .insert(relationship.relationship_id.as_str())
                        {
                            return Err(format!(
                                "Worksheet chart {} has invalid additional relationship metadata",
                                idx + 1
                            )
                            .into());
                        }
                        let (target, external) = match &relationship.target {
                            crate::xlsx::ChartRelationshipTarget::Embedded {
                                data,
                                content_type,
                                extension,
                            } => {
                                if content_type.is_empty()
                                    || extension.is_empty()
                                    || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                                {
                                    return Err(format!(
                                        "Worksheet chart {} has invalid embedded related resource",
                                        idx + 1
                                    )
                                    .into());
                                }
                                let resource_name = format!(
                                    "chartResource{}_{}_{}.{}",
                                    ws.sheet_id(),
                                    idx + 1,
                                    relationship_index + 1,
                                    extension.to_ascii_lowercase()
                                );
                                let resource_uri =
                                    PackURI::new(format!("/xl/chartResources/{resource_name}"))?;
                                chart_related_resources.push(BlobPart::new(
                                    resource_uri,
                                    content_type.clone(),
                                    data.clone(),
                                ));
                                (format!("../chartResources/{resource_name}"), false)
                            },
                            crate::xlsx::ChartRelationshipTarget::External { target } => {
                                if target.is_empty() {
                                    return Err(format!(
                                        "Worksheet chart {} has an empty external related target",
                                        idx + 1
                                    )
                                    .into());
                                }
                                (target.clone(), true)
                            },
                        };
                        chart_part.rels_mut().add_relationship(
                            relationship.relationship_type.clone(),
                            target,
                            relationship.relationship_id.clone(),
                            external,
                        );
                    }
                    let mut embedded_external_part = None;
                    let external_data_relationship_id = if let Some(external_data) =
                        chart.external_data_part.as_ref()
                    {
                        if !crate::xlsx::chart::is_chart_external_data_relationship_type(
                            &external_data.relationship_type,
                        ) {
                            return Err(format!(
                                    "Worksheet chart {} has invalid external-data relationship type '{}'",
                                    idx + 1,
                                    external_data.relationship_type
                                )
                                .into());
                        }
                        let (target, external) = match &external_data.target {
                            crate::xlsx::ChartExternalDataTarget::Embedded {
                                data,
                                content_type,
                                extension,
                            } => {
                                let expected_content_type =
                                    crate::xlsx::chart::chart_external_data_content_type(
                                        &external_data.relationship_type,
                                    )
                                    .expect("relationship type was validated above");
                                if content_type.is_empty()
                                    || content_type != expected_content_type
                                    || extension.is_empty()
                                    || !extension.bytes().all(|byte| byte.is_ascii_alphanumeric())
                                {
                                    return Err(format!(
                                            "Worksheet chart {} has invalid embedded external-data metadata",
                                            idx + 1
                                        )
                                        .into());
                                }
                                let external_name = format!(
                                    "chartData{}_{}.{}",
                                    ws.sheet_id(),
                                    idx + 1,
                                    extension.to_ascii_lowercase()
                                );
                                let external_uri =
                                    PackURI::new(format!("/xl/embeddings/{external_name}"))?;
                                embedded_external_part = Some(BlobPart::new(
                                    external_uri,
                                    content_type.clone(),
                                    data.clone(),
                                ));
                                (format!("../embeddings/{external_name}"), false)
                            },
                            crate::xlsx::ChartExternalDataTarget::Linked { target } => {
                                if target.is_empty() {
                                    return Err(format!(
                                            "Worksheet chart {} has an empty linked external-data target",
                                            idx + 1
                                        )
                                        .into());
                                }
                                (target.clone(), true)
                            },
                        };
                        let relationship_id = if let Some(relationship_id) = chart
                            .chart
                            .external_data
                            .as_ref()
                            .and_then(|metadata| metadata.relationship_id.as_deref())
                        {
                            if relationship_id.is_empty()
                                || chart_part.rels().get(relationship_id).is_some()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has a conflicting external-data relationship ID",
                                    idx + 1
                                )
                                .into());
                            }
                            chart_part.rels_mut().add_relationship(
                                external_data.relationship_type.clone(),
                                target,
                                relationship_id.to_string(),
                                external,
                            );
                            relationship_id.to_string()
                        } else if external {
                            chart_part.relate_to_ext(&target, &external_data.relationship_type)
                        } else {
                            chart_part.relate_to(&target, &external_data.relationship_type)
                        };
                        Some(relationship_id)
                    } else {
                        None
                    };

                    let mut user_shapes_part_to_add = None;
                    let mut user_shape_resources = Vec::new();
                    let user_shapes_relationship_id = if let Some(user_shapes) =
                        chart.user_shapes_part.as_ref()
                    {
                        let referenced_ids =
                            crate::xlsx::chart::chart_user_shapes_relationship_ids(
                                &user_shapes.xml,
                            )?;
                        let declared_ids: std::collections::HashSet<&str> = user_shapes
                            .relationships
                            .iter()
                            .map(|relationship| relationship.relationship_id.as_str())
                            .collect();
                        if declared_ids.len() != user_shapes.relationships.len()
                            || referenced_ids.len() != declared_ids.len()
                            || !referenced_ids
                                .iter()
                                .all(|id| declared_ids.contains(id.as_str()))
                        {
                            return Err(format!(
                                "Worksheet chart {} user-shapes relationship declarations do not match its XML",
                                idx + 1
                            )
                            .into());
                        }

                        let user_shapes_name =
                            format!("chartDrawing{}_{}.xml", ws.sheet_id(), idx + 1);
                        let user_shapes_uri =
                            PackURI::new(format!("/xl/drawings/{user_shapes_name}"))?;
                        let mut part = BlobPart::new(
                            user_shapes_uri,
                            ct::DML_CHARTSHAPES.to_string(),
                            user_shapes.xml.clone(),
                        );
                        for (relationship_index, relationship) in
                            user_shapes.relationships.iter().enumerate()
                        {
                            if relationship.relationship_id.is_empty()
                                || relationship.relationship_type.is_empty()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has invalid user-shapes relationship metadata",
                                    idx + 1
                                )
                                .into());
                            }
                            let (target, external) = match &relationship.target {
                                crate::xlsx::ChartUserShapesRelationshipTarget::Embedded {
                                    data,
                                    content_type,
                                    extension,
                                } => {
                                    if content_type.is_empty()
                                        || extension.is_empty()
                                        || !extension
                                            .bytes()
                                            .all(|byte| byte.is_ascii_alphanumeric())
                                    {
                                        return Err(format!(
                                            "Worksheet chart {} has invalid embedded user-shapes resource",
                                            idx + 1
                                        )
                                        .into());
                                    }
                                    let resource_name = format!(
                                        "chartShape{}_{}_{}.{}",
                                        ws.sheet_id(),
                                        idx + 1,
                                        relationship_index + 1,
                                        extension.to_ascii_lowercase()
                                    );
                                    let resource_uri =
                                        PackURI::new(format!("/xl/media/{resource_name}"))?;
                                    user_shape_resources.push(BlobPart::new(
                                        resource_uri,
                                        content_type.clone(),
                                        data.clone(),
                                    ));
                                    (format!("../media/{resource_name}"), false)
                                },
                                crate::xlsx::ChartUserShapesRelationshipTarget::External {
                                    target,
                                } => {
                                    if target.is_empty() {
                                        return Err(format!(
                                            "Worksheet chart {} has an empty external user-shapes target",
                                            idx + 1
                                        )
                                        .into());
                                    }
                                    (target.clone(), true)
                                },
                            };
                            part.rels_mut().add_relationship(
                                relationship.relationship_type.clone(),
                                target,
                                relationship.relationship_id.clone(),
                                external,
                            );
                        }
                        let target = format!("../drawings/{user_shapes_name}");
                        let relationship_id = if let Some(relationship_id) = chart
                            .chart
                            .user_shapes
                            .as_ref()
                            .and_then(|metadata| metadata.relationship_id.as_deref())
                        {
                            if relationship_id.is_empty()
                                || chart_part.rels().get(relationship_id).is_some()
                            {
                                return Err(format!(
                                    "Worksheet chart {} has a conflicting user-shapes relationship ID",
                                    idx + 1
                                )
                                .into());
                            }
                            chart_part.rels_mut().add_relationship(
                                rt::CHART_USER_SHAPES.to_string(),
                                target,
                                relationship_id.to_string(),
                                false,
                            );
                            relationship_id.to_string()
                        } else {
                            chart_part.relate_to(&target, rt::CHART_USER_SHAPES)
                        };
                        user_shapes_part_to_add = Some(part);
                        Some(relationship_id)
                    } else {
                        None
                    };

                    for relationship_id in
                        crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)?
                    {
                        if relationship_id.is_empty()
                            || chart_part.rels().get(&relationship_id).is_none()
                        {
                            return Err(format!(
                                "Worksheet chart {} fragment references missing relationship '{relationship_id}'",
                                idx + 1
                            )
                            .into());
                        }
                    }

                    let chart_xml = crate::xlsx::chart::generate_chart_xml_with_external_data_id(
                        chart_model,
                        external_data_relationship_id.as_deref(),
                        user_shapes_relationship_id.as_deref(),
                    )
                    .map_err(|e| format!("Failed to generate chart XML: {e}"))?;
                    chart_part.set_blob(chart_xml);
                    if let Some(external_part) = embedded_external_part {
                        self.package.add_part(Box::new(external_part));
                    }
                    for resource in chart_related_resources {
                        self.package.add_part(Box::new(resource));
                    }
                    for resource in user_shape_resources {
                        self.package.add_part(Box::new(resource));
                    }
                    if let Some(user_shapes_part) = user_shapes_part_to_add {
                        self.package.add_part(Box::new(user_shapes_part));
                    }
                    self.package.add_part(Box::new(chart_part));

                    // Add relationship from drawing to chart
                    drawing_part.relate_to(&format!("../charts/{chart_name}.xml"), rt::CHART);
                }

                self.package.add_part(Box::new(drawing_part));

                // Add relationship from worksheet to drawing
                Some(ws_part.relate_to(
                    &format!("../drawings/drawing{}.xml", ws.sheet_id()),
                    rt::DRAWING,
                ))
            } else {
                None
            };

            let mut pivot_table_rel_ids: Vec<String> = Vec::new();
            if let Some(targets) = pivot_table_targets_per_sheet.get(index) {
                for target in targets {
                    let rid = ws_part.relate_to(target, rt::PIVOT_TABLE);
                    pivot_table_rel_ids.push(rid);
                }
            }

            // Now generate worksheet XML with proper hyperlink relationship IDs and VML reference
            let ws_xml = ws.to_xml_with_part_rels(
                &mut data.shared_strings,
                &style_indices,
                crate::xlsx::writer::sheet::WorksheetPartRelationships {
                    hyperlinks: Some(&hyperlink_rel_ids),
                    vml_drawing: vml_rel_id.as_deref(),
                    pivot_tables: Some(&pivot_table_rel_ids),
                    tables: Some(&table_rel_ids),
                    drawing: drawing_rel_id.as_deref(),
                },
            )?;
            ws_part.set_blob(ws_xml.into_bytes());

            self.package.add_part(Box::new(ws_part));

            // Create relationship and track the ID (for ALL sheets)
            let rel_target = format!("worksheets/sheet{}.xml", ws.sheet_id());
            let rid = temp_wb_part.relate_to(
                &rel_target,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
            );
            worksheet_rel_ids.push(rid);
        }

        // Emit chartsheet parts: chartsheet XML, its drawing part with an
        // absolute-anchored graphic frame, and the hosted chart part.
        let mut chartsheet_rel_ids: Vec<String> = Vec::with_capacity(data.chart_sheets.len());
        for (chartsheet_index, chart_sheet) in data.chart_sheets.iter().enumerate() {
            let chartsheet_name = format!("sheet{}.xml", chartsheet_index + 1);
            let drawing_name = format!("drawingChartsheet{}.xml", chartsheet_index + 1);
            let chart_name = format!("chart{}_1", chart_sheet.sheet_id());

            // Chart part, with the same pivot-binding validation and
            // normalization applied to worksheet charts.
            let normalized_chart = Self::normalized_pivot_chart_model(
                &chart_sheet.chart().chart,
                chart_sheet.name(),
                &authored_pivot_tables,
            )?;
            let chart_model = normalized_chart
                .as_ref()
                .unwrap_or(&chart_sheet.chart().chart);
            let chart_xml = crate::xlsx::chart::generate_chart_xml_with_external_data_id(
                chart_model,
                None,
                None,
            )
            .map_err(|e| format!("Failed to generate chart XML: {e}"))?;
            let chart_uri = PackURI::new(format!("/xl/charts/{chart_name}.xml"))?;
            self.package.add_part(Box::new(BlobPart::new(
                chart_uri,
                ct::DML_CHART.to_string(),
                chart_xml,
            )));

            // Drawing part: one absolute anchor referencing the chart.
            let drawing_uri = PackURI::new(format!("/xl/drawings/{drawing_name}"))?;
            let mut drawing_part = BlobPart::new(
                drawing_uri,
                ct::OFC_DRAWING.to_string(),
                crate::xlsx::writer::chart_sheet::chart_sheet_drawing_xml(chart_sheet.name())
                    .into_bytes(),
            );
            drawing_part.relate_to(&format!("../charts/{chart_name}.xml"), rt::CHART);
            self.package.add_part(Box::new(drawing_part));

            // Chartsheet part referencing the drawing.
            let chartsheet_uri = PackURI::new(format!("/xl/chartsheets/{chartsheet_name}"))?;
            let mut chartsheet_part = BlobPart::new(
                chartsheet_uri,
                crate::xlsx::writer::chart_sheet::CHARTSHEET_CONTENT_TYPE.to_string(),
                crate::xlsx::writer::chart_sheet::chart_sheet_part_xml()?,
            );
            chartsheet_part.relate_to(&format!("../drawings/{drawing_name}"), rt::DRAWING);
            self.package.add_part(Box::new(chartsheet_part));

            chartsheet_rel_ids.push(temp_wb_part.relate_to(
                &format!("chartsheets/{chartsheet_name}"),
                crate::xlsx::writer::chart_sheet::CHARTSHEET_RELATIONSHIP_TYPE,
            ));
        }

        // Update shared strings
        let ss_xml = data.shared_strings.to_xml()?;
        let ss_uri = PackURI::new("/xl/sharedStrings.xml")?;
        let ss_part = BlobPart::new(
            ss_uri,
            ct::SML_SHARED_STRINGS.to_string(),
            ss_xml.into_bytes(),
        );
        self.package.add_part(Box::new(ss_part));

        // Create shared strings relationship
        temp_wb_part.relate_to(
            "sharedStrings.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
        );

        // Synchronize worksheet print settings with workbook-level defined names
        data.sync_print_settings_to_defined_names();

        // Now generate workbook XML with actual relationship IDs
        let external_reference_ids: Vec<String> = self
            .external_links
            .iter()
            .map(|link| link.relationship_id.clone())
            .collect();
        let sheet_entries: Vec<(String, u32, String)> = data
            .sheet_order()
            .iter()
            .map(|slot| match *slot {
                crate::xlsx::writer::workbook::SheetSlot::Worksheet(index) => {
                    let worksheet = &data.worksheets[index];
                    (
                        worksheet.name().to_string(),
                        worksheet.sheet_id(),
                        worksheet_rel_ids[index].clone(),
                    )
                },
                crate::xlsx::writer::workbook::SheetSlot::ChartSheet(index) => {
                    let chart_sheet = &data.chart_sheets[index];
                    (
                        chart_sheet.name().to_string(),
                        chart_sheet.sheet_id(),
                        chartsheet_rel_ids[index].clone(),
                    )
                },
            })
            .collect();
        let workbook_xml = data.generate_workbook_xml_ordered(
            &sheet_entries,
            &pivot_cache_rel_ids,
            &external_reference_ids,
        )?;
        temp_wb_part.set_blob(workbook_xml.into_bytes());

        // Add the workbook part to the package
        self.package.add_part(Box::new(temp_wb_part));

        Ok(())
    }

    /// Update the core.xml properties part.
    fn update_core_properties(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::part::BlobPart;

        let core_uri = PackURI::new("/docProps/core.xml")?;

        // Generate XML from properties
        let xml = self.properties.to_xml();

        // Create or update the core properties part
        let core_part = BlobPart::new(
            core_uri,
            ct::OPC_CORE_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.package.add_part(Box::new(core_part));

        Ok(())
    }

    /// Update the app.xml properties part with current worksheet information.
    fn update_app_properties(&mut self) -> SheetResult<()> {
        use litchi_opc::constants::content_type as ct;
        use litchi_opc::part::BlobPart;
        use std::fmt::Write;

        let app_uri = PackURI::new("/docProps/app.xml")?;

        // Get worksheet names from mutable_data if available, otherwise from package
        let worksheet_names: Vec<String> = if let Some(ref data) = self.mutable_data {
            data.worksheets
                .iter()
                .map(|ws| ws.name().to_string())
                .collect()
        } else {
            // Fallback to parsing from workbook.xml if no mutable data
            vec!["Sheet1".to_string()]
        };

        let worksheet_count = worksheet_names.len();

        // Generate app.xml XML
        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" "#);
        xml.push_str(
            r#"xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">"#,
        );
        xml.push_str("<Application>The Litchi Rust Library</Application>");
        xml.push_str("<DocSecurity>0</DocSecurity>");
        xml.push_str("<ScaleCrop>false</ScaleCrop>");

        // HeadingPairs: category name + count
        xml.push_str("<HeadingPairs>");
        xml.push_str(r#"<vt:vector size="2" baseType="variant">"#);
        xml.push_str("<vt:variant><vt:lpstr>Worksheet</vt:lpstr></vt:variant>");
        write!(
            xml,
            "<vt:variant><vt:i4>{}</vt:i4></vt:variant>",
            worksheet_count
        )
        .map_err(|e| format!("XML write error: {}", e))?;
        xml.push_str("</vt:vector>");
        xml.push_str("</HeadingPairs>");

        // TitlesOfParts: list of all worksheet names
        xml.push_str("<TitlesOfParts>");
        write!(
            xml,
            r#"<vt:vector size="{}" baseType="lpstr">"#,
            worksheet_count
        )
        .map_err(|e| format!("XML write error: {}", e))?;
        for name in &worksheet_names {
            // Escape XML special characters
            let escaped_name = name
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;")
                .replace('"', "&quot;")
                .replace('\'', "&apos;");
            write!(xml, "<vt:lpstr>{}</vt:lpstr>", escaped_name)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</vt:vector>");
        xml.push_str("</TitlesOfParts>");

        xml.push_str("<Company/>");
        xml.push_str("<LinksUpToDate>false</LinksUpToDate>");
        xml.push_str("<SharedDoc>false</SharedDoc>");
        xml.push_str("<HyperlinksChanged>false</HyperlinksChanged>");
        xml.push_str("<AppVersion>14.0000</AppVersion>");
        xml.push_str("</Properties>");

        // Create or update the app properties part
        let app_part = BlobPart::new(
            app_uri,
            ct::OFC_EXTENDED_PROPERTIES.to_string(),
            xml.into_bytes(),
        );

        self.package.add_part(Box::new(app_part));

        Ok(())
    }

    // ===== Workbook-level Features =====

    /// Hide a worksheet by index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.hide_sheet(0)?; // Hide the first sheet
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn hide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().hide_sheet(index)?;
        Ok(())
    }

    /// Unhide a worksheet by index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn unhide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().unhide_sheet(index)?;
        Ok(())
    }

    /// Check if a worksheet is hidden.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn is_sheet_hidden(&self, index: usize) -> bool {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.is_sheet_hidden(index))
            .unwrap_or(false)
    }

    /// Move a worksheet to a new position.
    ///
    /// # Arguments
    /// * `from_index` - Current worksheet index (0-based)
    /// * `to_index` - Target worksheet index (0-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.add_worksheet("Sheet3");
    /// wb.move_sheet(2, 0)?; // Move Sheet3 to the first position
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn move_sheet(&mut self, from_index: usize, to_index: usize) -> SheetResult<()> {
        if from_index >= self.worksheets.len() || to_index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .move_sheet(from_index, to_index)?;

        // Also update local worksheets vector
        let sheet = self.worksheets.remove(from_index);
        self.worksheets.insert(to_index, sheet);

        Ok(())
    }

    /// Set sheet visibility state.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    /// * `visibility` - Visibility state: "visible", "hidden", or "veryHidden"
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_sheet_visibility(0, "hidden")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_sheet_visibility(&mut self, index: usize, visibility: &str) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        if !matches!(visibility, "visible" | "hidden" | "veryHidden") {
            return Err(
                "Invalid visibility state. Must be 'visible', 'hidden', or 'veryHidden'".into(),
            );
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_sheet_visibility(index, visibility)?;
        Ok(())
    }

    /// Get sheet visibility state.
    ///
    /// Returns "visible", "hidden", or "veryHidden".
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn get_sheet_visibility(&self, index: usize) -> Option<&str> {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.get_sheet_visibility(index))
    }

    /// Set the active worksheet index.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based) to set as active
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.add_worksheet("Sheet2");
    /// wb.set_active_sheet(1)?; // Make Sheet2 active
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_active_sheet(&mut self, index: usize) -> SheetResult<()> {
        if index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        self.active_sheet_index = index;

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().set_active_sheet(index);
        Ok(())
    }

    /// Force formula recalculation when the workbook is opened.
    ///
    /// # Arguments
    /// * `force` - Whether to force recalculation
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_force_formula_recalculation(true);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_force_formula_recalculation(&mut self, force: bool) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_force_formula_recalculation(force);
    }

    /// Set the calculation mode for the workbook.
    ///
    /// # Arguments
    /// * `mode` - Calculation mode: "auto", "manual", or "autoNoTable"
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_calculation_mode("manual")?;
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_calculation_mode(&mut self, mode: &str) -> SheetResult<()> {
        if !matches!(mode, "auto" | "manual" | "autoNoTable") {
            return Err(
                "Invalid calculation mode. Must be 'auto', 'manual', or 'autoNoTable'".into(),
            );
        }

        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .set_calculation_mode(mode);
        Ok(())
    }

    /// Get the calculation mode for the workbook.
    ///
    /// Returns "auto", "manual", or "autoNoTable".
    pub fn get_calculation_mode(&self) -> &str {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.get_calculation_mode())
            .unwrap_or("auto")
    }

    /// Set the tab color for a worksheet.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    /// * `color` - RGB hex color (e.g., "FF0000" for red)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.set_tab_color(0, "FF0000")?; // Set red tab color
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_tab_color(&mut self, index: usize, color: &str) -> SheetResult<()> {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data
            .as_mut()
            .unwrap()
            .worksheet_mut(index)?
            .set_tab_color(color);
        Ok(())
    }

    /// Get the tab color for a worksheet.
    ///
    /// # Arguments
    /// * `index` - Worksheet index (0-based)
    pub fn get_tab_color(&self, index: usize) -> Option<&str> {
        self.mutable_data
            .as_ref()
            .and_then(|d| d.worksheets.get(index))
            .and_then(|ws| ws.tab_color())
    }

    /// Protect the workbook with optional password.
    ///
    /// # Arguments
    /// * `password` - Optional password (will be hashed)
    /// * `lock_structure` - Prevent adding/deleting sheets
    /// * `lock_windows` - Prevent resizing/moving workbook window
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// wb.protect_workbook(Some("password123"), true, false);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        lock_structure: bool,
        lock_windows: bool,
    ) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }

        self.mutable_data.as_mut().unwrap().protect_workbook(
            password,
            lock_structure,
            lock_windows,
        );
    }

    /// Replace the complete typed workbook-protection configuration.
    ///
    /// This supports workbook and revision locks with either legacy 16-bit
    /// verifiers or caller-supplied strong verifier metadata. Verifiers remain
    /// advisory and inert: this crate does not validate passwords or enforce
    /// an editing policy.
    pub fn set_workbook_protection(&mut self, protection: WorkbookProtectionMetadata) {
        if self.mutable_data.is_none() {
            self.mutable_data = Some(MutableWorkbookData::new());
        }
        self.mutable_data
            .as_mut()
            .unwrap()
            .set_workbook_protection(protection);
    }

    /// Unprotect the workbook.
    pub fn unprotect_workbook(&mut self) {
        if let Some(data) = self.mutable_data.as_mut() {
            data.unprotect_workbook();
        }
    }

    /// Check if the workbook is protected.
    pub fn is_workbook_protected(&self) -> bool {
        self.mutable_data.as_ref().is_some_and(|d| d.is_protected())
    }

    pub fn pivot_tables(&self) -> SheetResult<Vec<PivotTable>> {
        crate::xlsx::pivot::read_pivot_tables(self.package())
    }

    /// Load inert query-table metadata associated with a worksheet.
    pub fn query_tables_on_sheet(
        &self,
        sheet_name: &str,
    ) -> SheetResult<Vec<super::query_table::WorksheetQueryTable>> {
        let info = self
            .worksheets
            .iter()
            .find(|worksheet| worksheet.name == sheet_name)
            .cloned()
            .ok_or_else(|| format!("Worksheet '{sheet_name}' not found"))?;
        let mut worksheet = Worksheet::new(self, info);
        worksheet.load_data()?;
        Ok(worksheet.query_tables().to_vec())
    }

    pub fn pivot_tables_on_sheet(&self, sheet_name: &str) -> SheetResult<Vec<PivotTable>> {
        let all = self.pivot_tables()?;
        Ok(all
            .into_iter()
            .filter(|t| t.sheet_name == sheet_name)
            .collect())
    }

    /// Load the typed pivot-chart bindings anchored on one worksheet or
    /// chartsheet.
    ///
    /// Each returned pivot chart has its `c:pivotSource` name resolved to the
    /// typed pivot-table model; broken or dangling bindings are errors.
    /// Ordinary charts without a pivot source are excluded.
    pub fn pivot_charts_on_sheet(
        &self,
        sheet_name: &str,
    ) -> SheetResult<Vec<super::pivot_chart::PivotChart>> {
        Ok(super::pivot_chart::load_worksheet_pivot_charts(
            self.package(),
            sheet_name,
        )?)
    }

    /// Load the typed DrawingML shape and text-box inventory of one worksheet.
    ///
    /// Shapes, connection shapes, groups, and legacy OLE objects anchored on
    /// the worksheet's drawing part are returned in drawing order; pictures
    /// and charts are covered by `Worksheet::images()` and
    /// `Worksheet::charts()` instead. Everything is read-only and inert.
    pub fn shapes_on_sheet(
        &self,
        sheet_name: &str,
    ) -> SheetResult<super::shapes::XlsxWorksheetShapes> {
        Ok(super::shapes::load_worksheet_shapes(
            self.package(),
            sheet_name,
        )?)
    }

    // ===== Worksheet-level Writing Features =====
    // (These are mostly implemented via MutableWorksheet, exposed through worksheet_mut)

    // ============================================================================
    // Apache POI Features Implementation Status
    // ============================================================================
    //
    // ✅ FULLY IMPLEMENTED (Workbook-level):
    // - Hidden sheets: hide_sheet(), unhide_sheet(), is_sheet_hidden()
    // - Sheet ordering: move_sheet()
    // - Sheet visibility: set_sheet_visibility(), get_sheet_visibility()
    // - Active sheet: set_active_sheet()
    // - Workbook calculation mode: set_force_formula_recalculation(), set_calculation_mode(), get_calculation_mode()
    // - Named ranges: define_name(), define_name_local(), define_name_with_comment(), remove_name()
    // - Sheet tab color: set_tab_color(), get_tab_color()
    // - Workbook protection: protect_workbook(), unprotect_workbook(), is_workbook_protected()
    //
    // ✅ FULLY IMPLEMENTED (Worksheet reading - via Worksheet):
    // - Merged cells (reading): get_merged_regions(), is_merged_cell(), get_merge_region()
    // - Auto-filter (reading): get_auto_filter()
    // - Column width/Row height (reading): get_column_width(), get_row_height()
    // - Hyperlinks (reading): get_hyperlink(), get_hyperlinks()
    // - Comments (reading): get_cell_comment(), get_comments()
    // - Data validation (reading): get_data_validations()
    // - Conditional formatting (reading): get_conditional_formatting()
    // - Page setup (reading): get_page_setup()
    //
    // ✅ FULLY IMPLEMENTED (Worksheet writing - via MutableWorksheet):
    // - Cell values & formulas: set_cell_value(), set_cell_formula(), set_cell_formula_with_cache()
    // - Cell formatting: set_cell_format() with CellFormat (font, fill, border, number format)
    // - Merged cells: merge_cells()
    // - Column width/Row height: set_column_width(), set_row_height()
    // - Hide columns/rows: hide_column(), hide_row(), show_column(), show_row()
    // - Data validation: add_data_validation()
    // - Charts: add_chart() (basic support)
    // - Freeze panes: freeze_panes(), unfreeze_panes()
    // - Page setup: set_page_setup(), set_page_setup_with_options(), set_print_area(), clear_print_area()
    // - Auto-filter: set_auto_filter(), remove_auto_filter()
    // - Sheet protection: protect_sheet(), protect_sheet_with_options(), unprotect_sheet()
    // - Hyperlinks: set_hyperlink(), remove_hyperlink(), hyperlinks()
    // - Comments: set_cell_comment(), remove_comment(), comments()
    // - Conditional formatting: add_conditional_formatting(), clear_conditional_formatting()
    // - Row/column grouping: group_rows(), ungroup_rows(), group_columns(), ungroup_columns()
    //
    // ⚠️ BASIC IMPLEMENTATION (Data structures exist, XML generation would need enhancement):
    // - Hyperlinks: Stored but need relationship XML in worksheet rels
    // - Comments: Stored but need comments.xml part and VML drawing
    // - Conditional formatting: Stored but need full XML generation in worksheet
    // - Charts: Basic structure exists, needs DrawingML XML generation
    //
    // ⏳ NOT IMPLEMENTED (Advanced features requiring significant additional work):
    // - Pivot tables: add_pivot_table(), get_pivot_tables(), refresh_pivot_table()
    // - Images/Pictures: add_picture(), get_pictures(), delete_picture()
    // - Rich text in cells: set_rich_text_cell(), get_rich_text_cell()
    // - Subtotals: insert_subtotals(), remove_subtotals()
    // - Sparklines: add_sparkline(), get_sparklines()
    // - Slicers: add_slicer(), get_slicers()
    // - Timeline: add_timeline(), get_timelines()
    // - Power Query: get_power_query_connections()
    // - External links: get_external_links(), update_external_links()
    //
    // 📝 NOTES:
    // - Basic cell styling is fully supported via CellFormat (font, fill, border, number format)
    // - All reading operations work perfectly
    // - All core writing operations are implemented
    // - Advanced features like pivot tables, images would require substantial XML generation code
    // - The library is production-ready for standard Excel CRUD operations
}

fn validate_workbook_tables(data: &MutableWorkbookData) -> SheetResult<()> {
    use std::collections::HashSet;

    let table_count = data
        .worksheets
        .iter()
        .map(|worksheet| worksheet.tables().len())
        .sum();
    let mut ids = HashSet::with_capacity(table_count);
    let mut names = HashSet::with_capacity(table_count);
    let mut display_names = HashSet::with_capacity(table_count);
    for worksheet in &data.worksheets {
        for table in worksheet.tables() {
            if !ids.insert(table.id) {
                return Err(format!("duplicate workbook table ID {}", table.id).into());
            }
            if !names.insert(table.name.to_ascii_lowercase()) {
                return Err(format!("duplicate workbook table name '{}'", table.name).into());
            }
            if !display_names.insert(table.display_name.to_ascii_lowercase()) {
                return Err(format!(
                    "duplicate workbook table display name '{}'",
                    table.display_name
                )
                .into());
            }
        }
    }
    Ok(())
}

fn validate_threaded_comment_people<'a>(
    comments: impl IntoIterator<Item = &'a crate::xlsx::ThreadedComment>,
    person_list: Option<&crate::xlsx::PersonList>,
) -> SheetResult<()> {
    use std::collections::HashSet;

    let person_ids: HashSet<&str> = person_list
        .map(|people| {
            people
                .persons
                .iter()
                .map(|person| person.id.as_str())
                .collect()
        })
        .unwrap_or_default();
    for comment in comments {
        if !person_ids.contains(comment.person_id.as_str()) {
            return Err(format!(
                "threaded comment '{}' references missing person '{}'",
                comment.id, comment.person_id
            )
            .into());
        }
        for mention in &comment.mentions {
            if !person_ids.contains(mention.mention_person_id.as_str()) {
                return Err(format!(
                    "mention '{}' references missing person '{}'",
                    mention.mention_id, mention.mention_person_id
                )
                .into());
            }
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{Workbook, WorkbookProtectionMetadata, validate_threaded_comment_people};
    use crate::charts::{ChartExtensionList, ChartShapeProperties, plot_area::TypeGroup};
    use crate::xlsx::active_x::{
        ActiveXControlSet, ActiveXDescriptor, ActiveXProperty, LoadedActiveXControl, Persistence,
        WorksheetControl,
    };
    use litchi_core::sheet::{CellValue, WorkbookTrait, Worksheet as _};
    use litchi_opc::constants::{content_type as ct, relationship_type as rt};
    use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};

    use crate::xlsx::{
        ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship,
        ChartRelationshipTarget, ChartUserShapesPart, ChartUserShapesRelationship,
        ChartUserShapesRelationshipTarget, Mention, Person, PersonList, ProtectionPasswordVerifier,
        StrongProtectionPasswordVerifier, Table, TableColumn, ThreadedComment, WorksheetChart,
    };

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sales" sheetId="42" r:id="rId1"/></sheets>
  <definedNames>
    <definedName name="_xlnm.Print_Area" localSheetId="0">Sales!$A$1:$D$20</definedName>
    <definedName name="_xlnm.Print_Titles" localSheetId="0">Sales!$1:$2,Sales!$A:$B</definedName>
  </definedNames>
</workbook>"#;

    #[test]
    fn saves_and_reloads_complete_typed_workbook_protection() {
        let mut protection = WorkbookProtectionMetadata::new();
        protection.set_workbook_verifier(Some(ProtectionPasswordVerifier::Strong(
            StrongProtectionPasswordVerifier::new(
                "SHA-512",
                vec![1, 2, 3],
                vec![4, 5, 6],
                100_000,
            )
            .unwrap(),
        )));
        protection.set_revisions_verifier(Some(ProtectionPasswordVerifier::Legacy(0x00AF)));
        protection.set_structure_locked(true);
        protection.set_windows_locked(true);
        protection.set_revision_locked(true);

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("workbook-protection.xlsx");
        let mut workbook = Workbook::create().unwrap();
        workbook.set_workbook_protection(protection.clone());
        workbook.save(&path).unwrap();

        let reopened = Workbook::open(path).unwrap();
        assert_eq!(
            reopened.workbook_protection_metadata().unwrap(),
            Some(protection)
        );
    }

    #[test]
    fn saves_and_reloads_worksheet_tables_and_drawings() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("tables.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut table = Table::new(7, "SalesTable", "A1:B4");
        table.columns = vec![
            TableColumn::new(1, "Region"),
            TableColumn::new(2, "Revenue"),
        ];
        table.auto_filter_range = Some("A1:B4".to_string());
        let worksheet = workbook.worksheet_mut(0).unwrap();
        worksheet.add_table(table);
        worksheet
            .add_image(
                b"\x89PNG\r\n\x1a\nfixture".to_vec(),
                "png",
                1,
                1,
                4,
                2,
                Some("Logo"),
            )
            .unwrap();
        let mut worksheet_chart = WorksheetChart::bar_chart(
            "Revenue",
            "Sheet1!$A$2:$A$4",
            "Sheet1!$B$2:$B$4",
            ChartAnchor::new(3, 1, 9, 14),
        )
        .unwrap();
        worksheet_chart.chart.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId9"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        worksheet_chart.chart.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="shared"><x:reference r:id="rId1" r:link="rId10"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        worksheet.add_chart(
            worksheet_chart
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId9".to_string(),
                relationship_type: rt::IMAGE.to_string(),
                target: ChartRelationshipTarget::Embedded {
                    data: b"chart background".to_vec(),
                    content_type: ct::PNG.to_string(),
                    extension: "png".to_string(),
                },
            })
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId10".to_string(),
                relationship_type: rt::HYPERLINK.to_string(),
                target: ChartRelationshipTarget::External {
                    target: "https://example.test/chart".to_string(),
                },
            })
            .with_external_data_part(
                ChartExternalDataPart::embedded_workbook(b"PK chart workbook".to_vec()),
                Some(false),
            )
            .with_user_shapes_part(ChartUserShapesPart {
                xml: br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>1</cdr:y></cdr:to><cdr:pic><a:blip r:embed="rId5"/></cdr:pic></cdr:relSizeAnchor></c:userShapes>"#.to_vec(),
                relationships: vec![ChartUserShapesRelationship {
                    relationship_id: "rId5".to_string(),
                    relationship_type: rt::IMAGE.to_string(),
                    target: ChartUserShapesRelationshipTarget::Embedded {
                        data: b"shape image".to_vec(),
                        content_type: ct::PNG.to_string(),
                        extension: "png".to_string(),
                    },
                }],
            }),
        );

        workbook.save(&path).unwrap();
        let package = OpcPackage::open(&path).unwrap();
        let sheet_part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap();
        let sheet_xml = std::str::from_utf8(sheet_part.blob()).unwrap();
        assert!(sheet_xml.contains(r#"<drawing r:id="rId2"/>"#));
        let drawing_relationship = sheet_part.rels().get("rId2").unwrap();
        assert_eq!(drawing_relationship.reltype(), rt::DRAWING);
        assert_eq!(
            drawing_relationship.target_partname().unwrap().as_str(),
            "/xl/drawings/drawing1.xml"
        );
        let drawing_part = package
            .get_part(&drawing_relationship.target_partname().unwrap())
            .unwrap();
        assert_eq!(drawing_part.content_type(), ct::OFC_DRAWING);
        assert_eq!(
            drawing_part.rels().get("rId1").unwrap().reltype(),
            rt::IMAGE
        );
        assert_eq!(
            drawing_part.rels().get("rId2").unwrap().reltype(),
            rt::CHART
        );
        let chart_relationship = drawing_part.rels().get("rId2").unwrap();
        let chart_part = package
            .get_part(&chart_relationship.target_partname().unwrap())
            .unwrap();
        let external_relationship = chart_part.rels().get("rId1").unwrap();
        assert_eq!(external_relationship.reltype(), rt::PACKAGE);
        assert_eq!(
            package
                .get_part(&external_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"PK chart workbook"
        );
        let chart_xml = std::str::from_utf8(chart_part.blob()).unwrap();
        assert!(chart_xml.contains(r#"r:embed="rId9""#));
        let background_relationship = chart_part.rels().get("rId9").unwrap();
        assert_eq!(background_relationship.reltype(), rt::IMAGE);
        assert_eq!(
            package
                .get_part(&background_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"chart background"
        );
        let link_relationship = chart_part.rels().get("rId10").unwrap();
        assert_eq!(link_relationship.reltype(), rt::HYPERLINK);
        assert!(link_relationship.is_external());
        assert_eq!(link_relationship.target_ref(), "https://example.test/chart");
        assert!(chart_xml.contains(r#"<c:externalData r:id="rId1">"#));
        assert!(chart_xml.contains(r#"<c:autoUpdate val="0"/>"#));
        let user_shapes_relationship = chart_part.rels().get("rId2").unwrap();
        assert_eq!(user_shapes_relationship.reltype(), rt::CHART_USER_SHAPES);
        let user_shapes_part = package
            .get_part(&user_shapes_relationship.target_partname().unwrap())
            .unwrap();
        assert_eq!(user_shapes_part.content_type(), ct::DML_CHARTSHAPES);
        let shape_image_relationship = user_shapes_part.rels().get("rId5").unwrap();
        assert_eq!(shape_image_relationship.reltype(), rt::IMAGE);
        assert_eq!(
            package
                .get_part(&shape_image_relationship.target_partname().unwrap())
                .unwrap()
                .blob(),
            b"shape image"
        );

        let mut reopened = Workbook::open(&path).unwrap();
        let worksheet = reopened.get_worksheet(0).unwrap();

        assert_eq!(worksheet.tables().len(), 1);
        let table = &worksheet.tables()[0];
        assert_eq!(table.id, 7);
        assert_eq!(table.name, "SalesTable");
        assert_eq!(table.ref_range, "A1:B4");
        assert_eq!(table.column_names(), vec!["Region", "Revenue"]);
        assert_eq!(worksheet.images().len(), 1);
        assert_eq!(worksheet.images()[0].format, "png");
        assert_eq!(worksheet.images()[0].position, (0, 0, 3, 1));
        assert_eq!(worksheet.images()[0].description.as_deref(), Some("Logo"));
        assert_eq!(worksheet.charts().len(), 1);
        assert_eq!(worksheet.charts()[0].anchor.from_col, 3);
        assert_eq!(worksheet.charts()[0].anchor.to_row, 14);
        assert_eq!(
            worksheet.charts()[0]
                .chart
                .external_data
                .as_ref()
                .unwrap()
                .auto_update,
            Some(false)
        );
        let ChartExternalDataTarget::Embedded {
            data,
            content_type,
            extension,
        } = &worksheet.charts()[0]
            .external_data_part
            .as_ref()
            .unwrap()
            .target
        else {
            panic!("expected embedded chart workbook");
        };
        assert_eq!(data, b"PK chart workbook");
        assert_eq!(content_type, ct::OFC_PACKAGE);
        assert_eq!(extension, "xlsx");
        let user_shapes = worksheet.charts()[0].user_shapes_part.as_ref().unwrap();
        assert_eq!(user_shapes.relationships.len(), 1);
        assert_eq!(user_shapes.relationships[0].relationship_id, "rId5");
        let ChartUserShapesRelationshipTarget::Embedded { data, .. } =
            &user_shapes.relationships[0].target
        else {
            panic!("expected embedded chart user-shape resource");
        };
        assert_eq!(data, b"shape image");
        assert_eq!(worksheet.charts()[0].additional_relationships.len(), 2);
        let background = worksheet.charts()[0]
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId9")
            .unwrap();
        let ChartRelationshipTarget::Embedded { data, .. } = &background.target else {
            panic!("expected embedded chart relationship resource");
        };
        assert_eq!(data, b"chart background");
        let link = worksheet.charts()[0]
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId10")
            .unwrap();
        let ChartRelationshipTarget::External { target } = &link.target else {
            panic!("expected external chart relationship target");
        };
        assert_eq!(target, "https://example.test/chart");
        let TypeGroup::Bar(group) = &worksheet.charts()[0].chart.plot_area.type_groups[0] else {
            panic!("expected reopened bar chart");
        };
        let series = &group.common.series[0];
        assert_eq!(
            series
                .categories
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$A$2:$A$4"
        );
        assert_eq!(
            series
                .values
                .as_ref()
                .unwrap()
                .source_ref
                .as_ref()
                .unwrap()
                .formula,
            "Sheet1!$B$2:$B$4"
        );

        let second_path = directory.path().join("tables-roundtrip.xlsx");
        reopened.save(&second_path).unwrap();
        let second_reopen = Workbook::open(&second_path).unwrap();
        assert_eq!(
            second_reopen.get_worksheet(0).unwrap().charts()[0]
                .chart
                .external_data
                .as_ref()
                .unwrap()
                .relationship_id
                .as_deref(),
            Some("rId1")
        );
    }

    #[test]
    fn rejects_dangling_chart_fragment_relationships() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("dangling-chart-relationship.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut chart = WorksheetChart::bar_chart(
            "Revenue",
            "Sheet1!$A$2:$A$4",
            "Sheet1!$B$2:$B$4",
            ChartAnchor::new(3, 1, 9, 14),
        )
        .unwrap();
        let TypeGroup::Bar(group) = &mut chart.chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        let series = &mut group.common.series[0];
        series.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId403"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut point = crate::charts::DataPoint::new(0);
        point.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="dangling"><x:reference r:id="rId404"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        series.data_points.push(point);
        chart.chart.chart_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="chart"><x:reference r:id="rId405"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId406"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.type_groups[0]
            .common_mut()
            .extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="group"><x:reference r:id="rId407"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        chart.chart.plot_area.data_table = Some(crate::charts::DataTable {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId408"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::DataTable::default()
        });
        let TypeGroup::Bar(group) = &mut chart.chart.plot_area.type_groups[0] else {
            panic!("expected a bar chart");
        };
        group.series_lines.push(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId409"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        let mut line =
            crate::charts::LineTypeGroup::new(crate::charts::types::BarGrouping::Standard);
        line.up_down_bars = Some(crate::charts::UpDownBars {
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="bars"><x:reference r:id="rId410"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::UpDownBars::default()
        });
        let mut line_series = crate::charts::Series::new(0);
        line_series.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId424"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="series-marker"><x:reference r:id="rId425"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut line_point = crate::charts::DataPoint::new(0);
        line_point.marker_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId426"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_point.marker_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="point-marker"><x:reference r:id="rId427"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.data_points.push(line_point);
        let mut line_labels = crate::charts::DataLabels::new();
        line_labels.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId430"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId431"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.leader_lines = Some(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId432"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        line_labels.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="labels"><x:reference r:id="rId433"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let mut line_label = crate::charts::DataLabel::new(0);
        line_label.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId434"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_label.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId435"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_label.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="label"><x:reference r:id="rId436"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_labels.labels.push(line_label);
        line_series.data_labels = Some(line_labels);
        line_series.error_bars.push(crate::charts::series::ErrorBar {
            direction: crate::charts::series::ErrorBarDirection::Y,
            error_type: crate::charts::series::ErrorBarType::Both,
            value_type: crate::charts::series::ErrorBarValueType::Fixed,
            value: Some(1.0),
            plus_values: None,
            minus_values: None,
            no_end_cap: false,
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId437"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="error-bars"><x:reference r:id="rId438"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        let mut line_trendline = crate::charts::series::Trendline::linear();
        line_trendline.show_label = true;
        line_trendline.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId439"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId440"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId441"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.label_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="trendline-label"><x:reference r:id="rId442"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_trendline.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="trendline"><x:reference r:id="rId443"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        line_series.trendlines.push(line_trendline);
        line.common.series.push(line_series);
        chart
            .chart
            .plot_area
            .type_groups
            .push(TypeGroup::Line(line));
        let mut surface = crate::charts::plot_area::SurfaceTypeGroup::new();
        let mut band = crate::charts::plot_area::BandFormat::new(0);
        band.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId444"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        surface.band_formats = Some(vec![band]);
        chart
            .chart
            .plot_area
            .type_groups
            .push(TypeGroup::Surface(surface));
        let mut legend = crate::charts::Legend {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId411"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::Legend::default()
        };
        let mut legend_entry = crate::charts::legend::LegendEntry::new(0);
        legend_entry.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="entry"><x:reference r:id="rId412"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        legend.entries.push(legend_entry);
        chart.chart.legend = Some(legend);
        chart.chart.title_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="title"><x:reference r:id="rId413"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let axis_common = chart.chart.plot_area.axes[0].common_mut();
        axis_common.title = Some(crate::charts::TitleText::from_string("Axis"));
        axis_common.title_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId414"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        axis_common.major_gridlines = Some(crate::charts::ChartLines {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId415"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
        });
        axis_common.scaling_extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="scaling"><x:reference r:id="rId416"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        axis_common.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="axis"><x:reference r:id="rId417"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        let crate::charts::Axis::Value(value_axis) = &mut chart.chart.plot_area.axes[1] else {
            panic!("expected value axis");
        };
        let mut display_units =
            crate::charts::axis::DisplayUnits::built_in(crate::charts::axis::BuiltInUnit::Millions);
        display_units.show_label = true;
        display_units.label_shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId418"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        display_units.label_text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId419"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        display_units.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="display-units"><x:reference r:id="rId420"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        value_axis.display_units = Some(Box::new(display_units));
        let mut pivot_format = crate::charts::PivotFormat::new(0);
        pivot_format.shape_properties = Some(
            ChartShapeProperties::from_xml(
                br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId421"/></a:blipFill></c:spPr>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.text_properties = Some(
            crate::charts::ChartTextProperties::from_xml(
                br#"<c:txPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:hlinkClick r:id="rId422"/></a:defRPr></a:pPr></a:p></c:txPr>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.extension_list = Some(
            ChartExtensionList::from_xml(
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="pivot"><x:reference r:id="rId423"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        pivot_format.marker = Some(crate::charts::Marker {
            shape_properties: Some(
                ChartShapeProperties::from_xml(
                    br#"<c:spPr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blipFill><a:blip r:embed="rId428"/></a:blipFill></c:spPr>"#.to_vec(),
                )
                .unwrap(),
            ),
            extension_list: Some(
                ChartExtensionList::from_xml(
                    br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="pivot-marker"><x:reference r:id="rId429"/></c:ext></c:extLst>"#.to_vec(),
                )
                .unwrap(),
            ),
            ..crate::charts::Marker::default()
        });
        chart.chart.pivot_formats = Some(vec![pivot_format]);
        let fragment_ids = crate::xlsx::chart::chart_fragment_relationship_ids(&chart.chart)
            .expect("relationship-bearing fragments should be valid XML");
        assert_eq!(
            fragment_ids,
            [
                "rId403".to_string(),
                "rId404".to_string(),
                "rId405".to_string(),
                "rId406".to_string(),
                "rId407".to_string(),
                "rId408".to_string(),
                "rId409".to_string(),
                "rId410".to_string(),
                "rId411".to_string(),
                "rId412".to_string(),
                "rId413".to_string(),
                "rId414".to_string(),
                "rId415".to_string(),
                "rId416".to_string(),
                "rId417".to_string(),
                "rId418".to_string(),
                "rId419".to_string(),
                "rId420".to_string(),
                "rId421".to_string(),
                "rId422".to_string(),
                "rId423".to_string(),
                "rId424".to_string(),
                "rId425".to_string(),
                "rId426".to_string(),
                "rId427".to_string(),
                "rId428".to_string(),
                "rId429".to_string(),
                "rId430".to_string(),
                "rId431".to_string(),
                "rId432".to_string(),
                "rId433".to_string(),
                "rId434".to_string(),
                "rId435".to_string(),
                "rId436".to_string(),
                "rId437".to_string(),
                "rId438".to_string(),
                "rId439".to_string(),
                "rId440".to_string(),
                "rId441".to_string(),
                "rId442".to_string(),
                "rId443".to_string(),
                "rId444".to_string(),
            ]
            .into()
        );
        workbook.worksheet_mut(0).unwrap().add_chart(chart);

        let error = workbook.save(&path).unwrap_err().to_string();
        assert!(error.contains("fragment references missing relationship"));
        assert!(
            [
                "rId403", "rId404", "rId405", "rId406", "rId407", "rId408", "rId409", "rId410",
                "rId411", "rId412", "rId413", "rId414", "rId415", "rId416", "rId417", "rId418",
                "rId419", "rId420", "rId421", "rId422", "rId423", "rId424", "rId425", "rId426",
                "rId427", "rId428", "rId429", "rId430", "rId431", "rId432", "rId433", "rId434",
                "rId435", "rId436", "rId437", "rId438", "rId439", "rId440", "rId441", "rId442",
                "rId443", "rId444",
            ]
            .iter()
            .any(|relationship_id| error.contains(relationship_id))
        );
    }

    #[test]
    fn rejects_duplicate_table_identity_across_worksheets() {
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("duplicate-tables.xlsx");
        let mut workbook = Workbook::create().unwrap();
        let mut first = Table::new(1, "Sales", "A1:A2");
        first.columns.push(TableColumn::new(1, "Value"));
        workbook.worksheet_mut(0).unwrap().add_table(first);

        let worksheet = workbook.add_worksheet("Other");
        let mut duplicate = Table::new(1, "OtherTable", "A1:A2");
        duplicate.columns.push(TableColumn::new(1, "Value"));
        worksheet.add_table(duplicate);

        let error = workbook.save(path).unwrap_err().to_string();
        assert!(error.contains("duplicate workbook table ID 1"));
    }

    #[test]
    fn validates_threaded_comment_people_references() {
        let person_id = "{11111111-1111-1111-1111-111111111111}";
        let missing_id = "{22222222-2222-2222-2222-222222222222}";
        let people = PersonList {
            persons: vec![Person {
                id: person_id.into(),
                ..Default::default()
            }],
        };
        let comment = ThreadedComment {
            id: "{33333333-3333-3333-3333-333333333333}".into(),
            person_id: person_id.into(),
            mentions: vec![Mention {
                mention_id: "{44444444-4444-4444-4444-444444444444}".into(),
                mention_person_id: person_id.into(),
                ..Default::default()
            }],
            ..Default::default()
        };

        validate_threaded_comment_people([&comment], Some(&people)).unwrap();
        assert!(validate_threaded_comment_people([&comment], None).is_err());

        let mut invalid = comment;
        invalid.mentions[0].mention_person_id = missing_id.into();
        assert!(validate_threaded_comment_people([&invalid], Some(&people)).is_err());
    }

    #[test]
    fn parses_quoted_and_bounded_print_defined_names() {
        assert_eq!(
            Workbook::parse_print_area("'Sales, West'!$A$1:$D$20,'Other'!$A$1"),
            Some("$A$1:$D$20".to_string())
        );
        assert_eq!(
            Workbook::parse_print_titles("'O''Brien, West'!$1:$2,'O''Brien, West'!$A:$B"),
            (Some("$1:$2".to_string()), Some("$A:$B".to_string()))
        );
        assert_eq!(Workbook::parse_print_area("Sales!XFE1"), None);
        assert_eq!(
            Workbook::parse_print_titles("Sales!$0:$1,Sales!$A:$XFE"),
            (None, None)
        );
    }

    const WORKSHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0" showGridLines="1"/>
    <sheetView workbookViewId="2" view="pageBreakPreview" topLeftCell="B2"
               zoomScale="90" zoomScaleNormal="110" showOutlineSymbols="0">
      <pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>
      <selection pane="bottomRight" activeCell="B3" sqref="B3:C4"/>
    </sheetView>
  </sheetViews>
  <cols><col min="2" max="3" width="12.5" hidden="1" customWidth="1"/></cols>
  <sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>
  <autoFilter ref="A1:E10">
    <sortState ref="A2:E10" caseSensitive="1" sortMethod="stroke">
      <sortCondition ref="E2:E10" descending="true" sortBy="value"/>
    </sortState>
  </autoFilter>
  <mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>
  <hyperlinks>
    <hyperlink ref="B2:C3" r:id="rId1" location="Section 1"
               display="Example &amp; Co" tooltip="Open example"/>
    <hyperlink ref="D4" location="&apos;Other Sheet&apos;!A1"/>
  </hyperlinks>
  <dataValidations count="1">
    <dataValidation sqref="E1:E2" type="whole" operator="between" allowBlank="1"
                    showErrorMessage="true" errorTitle="Out of range">
      <formula1>1</formula1><formula2>10</formula2>
    </dataValidation>
  </dataValidations>
  <conditionalFormatting sqref="E1:E2">
    <cfRule type="expression" priority="1"><formula>E1&gt;0</formula></cfRule>
  </conditionalFormatting>
  <pageSetup paperSize="9" orientation="landscape" scale="110"
             fitToWidth="1" fitToHeight="2"/>
  <rowBreaks count="1" manualBreakCount="1">
    <brk id="20" min="0" max="16383" man="1" pt="true"/>
  </rowBreaks>
  <colBreaks count="1" manualBreakCount="0"><brk id="3"/></colBreaks>
  <extLst>
    <ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"
         xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
      <x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">
        <x14:sparklineGroup type="stacked" dateAxis="1">
          <x14:colorSeries rgb="FF123456"/>
          <x14:sparklines>
            <x14:sparkline><xm:f>Sales!A1:A3</xm:f><xm:sqref>F2</xm:sqref></x14:sparkline>
            <x14:sparkline><xm:f>Sales!B1:B3</xm:f><xm:sqref>G2</xm:sqref></x14:sparkline>
          </x14:sparklines>
        </x14:sparklineGroup>
      </x14:sparklineGroups>
    </ext>
  </extLst>
</worksheet>"#;

    const COMMENTS_XML: &str = r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <authors><author>Alice &amp; Bob</author></authors>
  <commentList>
    <comment ref="C3" authorId="0" guid="{comment-guid}" shapeId="5">
      <text><r><t>Hello </t></r><r><t>world</t></r></text>
    </comment>
  </commentList>
</comments>"#;

    fn package_with_worksheet_relationship(reltype: &str, external: bool) -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            WORKBOOK_XML.as_bytes().to_vec(),
        );

        let relationship_id = if external {
            workbook_part.relate_to_ext("https://example.invalid/sheet.xml", reltype)
        } else {
            workbook_part.relate_to("custom/sales-data.xml", reltype)
        };
        assert_eq!(relationship_id, "rId1");
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));

        if !external {
            let mut worksheet_part = BlobPart::new(
                PackURI::new("/xl/custom/sales-data.xml").unwrap(),
                ct::SML_WORKSHEET.to_string(),
                WORKSHEET_XML.as_bytes().to_vec(),
            );
            let hyperlink_reltype = if reltype == rt::STRICT_WORKSHEET {
                rt::STRICT_HYPERLINK
            } else {
                rt::HYPERLINK
            };
            assert_eq!(
                worksheet_part.relate_to_ext("https://example.com/report", hyperlink_reltype),
                "rId1"
            );
            let comments_reltype = if reltype == rt::STRICT_WORKSHEET {
                rt::STRICT_COMMENTS
            } else {
                rt::COMMENTS
            };
            assert_eq!(
                worksheet_part.relate_to("../comments/custom-comments.xml", comments_reltype),
                "rId2"
            );
            package.add_part(Box::new(worksheet_part));
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/comments/custom-comments.xml").unwrap(),
                ct::SML_COMMENTS.to_string(),
                COMMENTS_XML.as_bytes().to_vec(),
            )));
        }

        package
    }

    #[test]
    fn exposes_and_transactionally_removes_worksheet_web_extension_bindings() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        package.get_part_mut(&worksheet_uri).unwrap().set_blob(
            include_bytes!("../../../../test-data/ooxml/web_extensions/worksheet_bindings.xml")
                .to_vec(),
        );
        let mut workbook = Workbook::new(package).unwrap();

        let bindings = workbook.worksheet_web_extension_bindings(0).unwrap();
        assert_eq!(bindings.len(), 2);
        assert_eq!(bindings[0].application_reference(), "sales-table");
        let worksheet = workbook.get_worksheet(0).unwrap();
        assert_eq!(worksheet.web_extension_bindings(), bindings);

        workbook.remove_worksheet_web_extension_bindings(0).unwrap();
        assert!(
            workbook
                .worksheet_web_extension_bindings(0)
                .unwrap()
                .is_empty()
        );

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("bindings-removed.xlsx");
        workbook.save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        assert!(
            reopened
                .worksheet_web_extension_bindings(0)
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn preserves_worksheet_web_extension_bindings_during_materialization() {
        let mut workbook = Workbook::create().unwrap();
        workbook
            .package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
            .unwrap()
            .set_blob(
                include_bytes!("../../../../test-data/ooxml/web_extensions/worksheet_bindings.xml")
                    .to_vec(),
            );
        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value(1, 1, "materialized");

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("bindings-preserved.xlsx");
        workbook.save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        let bindings = reopened.worksheet_web_extension_bindings(0).unwrap();
        assert_eq!(bindings.len(), 2);
        assert_eq!(bindings[1].application_reference(), "sales-point");
    }

    #[test]
    fn preserves_inert_active_x_graph_during_materialization() {
        let mut workbook = Workbook::create().unwrap();
        let value = ActiveXControlSet {
            controls: vec![LoadedActiveXControl {
                control: WorksheetControl {
                    shape_id: 17,
                    relationship_id: "rIdGeneratedControl".into(),
                    name: Some("Generated".into()),
                    properties: None,
                },
                descriptor_uri: PackURI::new("/xl/activeX/generated.xml").unwrap(),
                descriptor: ActiveXDescriptor {
                    class_id: "{00000000-0000-0000-0000-000000000000}".into(),
                    license: None,
                    persistence: Persistence::PropertyBag,
                    relationship_id: None,
                    properties: vec![ActiveXProperty {
                        name: "Caption".into(),
                        value: Some("inert".into()),
                        object: None,
                    }],
                },
                binaries: Vec::new(),
                preview: None,
            }],
        };
        workbook
            .store_worksheet_active_x_controls(0, &value)
            .unwrap();
        workbook
            .worksheet_mut(0)
            .unwrap()
            .set_cell_value(1, 1, "materialized");

        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("active-x-preserved.xlsx");
        workbook.save(&path).unwrap();
        let reopened = Workbook::open(&path).unwrap();
        let loaded = reopened.worksheet_active_x_controls(0).unwrap();
        assert_eq!(loaded.controls.len(), 1);
        assert_eq!(loaded.controls[0].control.shape_id, 17);
        assert_eq!(loaded.controls[0].descriptor, value.controls[0].descriptor);
    }

    fn package_with_custom_workbook_parts() -> OpcPackage {
        let mut package = OpcPackage::new();
        let workbook_uri = PackURI::new("/custom/book/main.xml").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            ct::SML_SHEET_MAIN.to_string(),
            br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                    xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                    <sheets><sheet name="Custom" sheetId="1" r:id="rId1"/></sheets>
                </workbook>"#
                .to_vec(),
        );
        assert_eq!(
            workbook_part.relate_to("../sheets/data.xml", rt::STRICT_WORKSHEET),
            "rId1"
        );
        assert_eq!(
            workbook_part.relate_to("../assets/strings.xml", rt::STRICT_SHARED_STRINGS),
            "rId2"
        );
        assert_eq!(
            workbook_part.relate_to("../assets/styles.xml", rt::STRICT_STYLES),
            "rId3"
        );
        package.relate_to("custom/book/main.xml", rt::STRICT_OFFICE_DOCUMENT);
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/sheets/data.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main">
                    <sheetData><row r="1"><c r="A1" t="s" s="0"><v>0</v></c></row></sheetData>
                </worksheet>"#
                .to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/strings.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_string(),
            br#"<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="1" uniqueCount="1">
                    <si><t>Relationship resolved</t></si>
                </sst>"#
                .to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/styles.xml").unwrap(),
            ct::SML_STYLES.to_string(),
            crate::xlsx::template::default_styles_xml()
                .as_bytes()
                .to_vec(),
        )));
        package
    }

    fn replace_hyperlink_relationship(package: &mut OpcPackage, reltype: &str, external: bool) {
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
        relationships.remove("rId1").unwrap();
        relationships.add_relationship(
            reltype.to_string(),
            "https://example.com/replaced".to_string(),
            "rId1".to_string(),
            external,
        );
    }

    #[test]
    fn loads_worksheet_from_relationship_target() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::WORKSHEET, false)).unwrap();
        let worksheet = workbook.get_worksheet(0).unwrap();

        assert_eq!(
            worksheet.cell_value(1, 1).unwrap().as_ref(),
            &CellValue::Int(7)
        );
        assert_eq!(worksheet.get_column_width(2), Some(12.5));
        assert!(worksheet.is_column_hidden(3));
        assert_eq!(worksheet.get_merged_regions(), &[(2, 2, 3, 3)]);
        let external_link = worksheet.get_hyperlink(3, 3).unwrap();
        assert_eq!(external_link.target, "https://example.com/report");
        assert_eq!(external_link.location.as_deref(), Some("Section 1"));
        assert_eq!(external_link.display.as_deref(), Some("Example & Co"));
        assert_eq!(external_link.tooltip.as_deref(), Some("Open example"));
        assert_eq!(
            worksheet.get_hyperlink(4, 4).unwrap().target,
            "'Other Sheet'!A1"
        );
        let validation = &worksheet.get_data_validations()[0];
        assert_eq!(validation.range, "E1:E2");
        assert_eq!(validation.operator.as_deref(), Some("between"));
        assert_eq!(validation.formula.as_deref(), Some("1"));
        assert_eq!(validation.formula2.as_deref(), Some("10"));
        assert!(validation.allow_blank);
        assert!(validation.show_error_message);
        assert_eq!(validation.error_title.as_deref(), Some("Out of range"));
        assert_eq!(worksheet.get_conditional_formatting().len(), 1);
        assert_eq!(
            worksheet.get_conditional_formatting()[0].rule_type,
            "expression"
        );
        assert_eq!(worksheet.get_page_setup().paper_size, Some(9));
        assert!(worksheet.get_page_setup().landscape);
        assert_eq!(worksheet.get_page_setup().scale, Some(110));
        assert_eq!(worksheet.row_breaks().len(), 1);
        assert!(worksheet.row_breaks()[0].manual);
        assert!(worksheet.row_breaks()[0].pivot);
        assert_eq!(worksheet.col_breaks()[0].max, 1_048_575);
        assert!(!worksheet.col_breaks()[0].manual);
        let filter = worksheet.get_auto_filter().unwrap();
        assert_eq!(filter.range.as_deref(), Some("A1:E10"));
        let sort = filter.sort_state.as_ref().unwrap();
        assert_eq!(sort.ref_range, "A2:E10");
        assert_eq!(sort.case_sensitive, Some(true));
        assert_eq!(sort.sort_method, Some(crate::xlsx::SortMethod::Stroke));
        assert_eq!(sort.conditions.len(), 1);
        assert_eq!(sort.conditions[0].descending, Some(true));
        assert_eq!(worksheet.sheet_views().len(), 2);
        let view = worksheet.sheet_view().unwrap();
        assert_eq!(view.workbook_view_id, Some(2));
        assert_eq!(
            view.view_type,
            Some(crate::xlsx::SheetViewType::PageBreakPreview)
        );
        assert_eq!(view.top_left_cell.as_deref(), Some("B2"));
        assert_eq!(view.zoom_scale, Some(90));
        assert_eq!(view.zoom_scale_normal, Some(110));
        assert_eq!(view.show_outline_symbols, Some(false));
        assert_eq!(view.pane.as_ref().unwrap().x_split, Some(1.0));
        assert_eq!(
            view.pane.as_ref().unwrap().state,
            Some(crate::xlsx::SheetPaneState::Frozen)
        );
        assert_eq!(view.selections[0].sqref.as_deref(), Some("B3:C4"));
        let sparkline = &worksheet.sparkline_groups()[0];
        assert_eq!(
            sparkline.sparkline_type,
            crate::xlsx::SparklineType::WinLoss
        );
        assert!(sparkline.options.date_axis);
        assert_eq!(sparkline.sparklines.len(), 2);
        assert_eq!(sparkline.sparklines[0].location, "F2");
        assert_eq!(worksheet.get_print_area(), Some("$A$1:$D$20"));
        assert_eq!(worksheet.get_repeating_rows(), Some("$1:$2"));
        assert_eq!(worksheet.get_repeating_columns(), Some("$A:$B"));
        let comment = worksheet.get_cell_comment(3, 3).unwrap();
        assert_eq!(comment.author.as_deref(), Some("Alice & Bob"));
        assert_eq!(comment.author_id, 0);
        assert_eq!(comment.text, "Hello world");
        assert_eq!(comment.guid.as_deref(), Some("{comment-guid}"));
        assert_eq!(comment.shape_id, Some(5));
    }

    #[test]
    fn resolves_strict_custom_workbook_related_parts() {
        let workbook = Workbook::new(package_with_custom_workbook_parts()).unwrap();
        let worksheet = workbook.get_worksheet(0).unwrap();

        assert_eq!(worksheet.name(), "Custom");
        assert_eq!(
            worksheet.cell_value(1, 1).unwrap().as_str(),
            Some("Relationship resolved")
        );
    }

    #[test]
    fn rejects_invalid_workbook_related_parts() {
        let mut package = package_with_custom_workbook_parts();
        let workbook_uri = PackURI::new("/custom/book/main.xml").unwrap();
        package
            .get_part_mut(&workbook_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::STRICT_SHARED_STRINGS.to_string(),
                "../assets/other-strings.xml".to_string(),
                "rId4".to_string(),
                false,
            );
        assert!(Workbook::new(package).is_err());

        let mut package = package_with_custom_workbook_parts();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/assets/styles.xml").unwrap(),
            ct::SML_SHARED_STRINGS.to_string(),
            Vec::new(),
        )));
        assert!(Workbook::new(package).is_err());
    }

    #[test]
    fn accepts_strict_worksheet_relationship_type() {
        let workbook = Workbook::new(package_with_worksheet_relationship(
            rt::STRICT_WORKSHEET,
            false,
        ))
        .unwrap();

        assert_eq!(workbook.worksheet_by_index(0).unwrap().row_count(), 1);
    }

    #[test]
    fn rejects_external_comments_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
        relationships.remove("rId2").unwrap();
        relationships.add_relationship(
            rt::COMMENTS.to_string(),
            "https://example.com/comments.xml".to_string(),
            "rId2".to_string(),
            true,
        );
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_duplicate_comments_relationships() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        package
            .get_part_mut(&worksheet_uri)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::COMMENTS.to_string(),
                "../comments/other-comments.xml".to_string(),
                "rId3".to_string(),
                false,
            );
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_wrong_comments_content_type() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/comments/custom-comments.xml").unwrap(),
            ct::SML_WORKSHEET.to_string(),
            COMMENTS_XML.as_bytes().to_vec(),
        )));
        let workbook = Workbook::new(package).unwrap();

        assert!(workbook.get_worksheet(0).is_err());
    }

    #[test]
    fn rejects_non_worksheet_relationship_type() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::CHART, false)).unwrap();
        let error = workbook
            .worksheet_by_index(0)
            .err()
            .expect("non-worksheet relationship must fail")
            .to_string();

        assert!(error.contains("invalid type"), "unexpected error: {error}");
    }

    #[test]
    fn rejects_external_worksheet_target() {
        let workbook =
            Workbook::new(package_with_worksheet_relationship(rt::WORKSHEET, true)).unwrap();
        let error = workbook
            .worksheet_by_index(0)
            .err()
            .expect("external worksheet relationship must fail")
            .to_string();

        assert!(
            error.contains("external target"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn rejects_missing_hyperlink_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        let worksheet_uri = PackURI::new("/xl/custom/sales-data.xml").unwrap();
        let worksheet_part = package.get_part_mut(&worksheet_uri).unwrap();
        let xml = std::str::from_utf8(worksheet_part.blob())
            .unwrap()
            .replace("r:id=\"rId1\"", "r:id=\"missingLink\"");
        worksheet_part.set_blob(xml.into_bytes());
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("missing hyperlink relationship must fail")
            .to_string();

        assert!(
            error.contains("missing relationship 'missingLink'"),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn rejects_invalid_hyperlink_relationship_type() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        replace_hyperlink_relationship(&mut package, rt::CHART, true);
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("non-hyperlink relationship must fail")
            .to_string();

        assert!(error.contains("invalid type"), "unexpected error: {error}");
    }

    #[test]
    fn rejects_internal_hyperlink_relationship() {
        let mut package = package_with_worksheet_relationship(rt::WORKSHEET, false);
        replace_hyperlink_relationship(&mut package, rt::HYPERLINK, false);
        let workbook = Workbook::new(package).unwrap();
        let error = workbook
            .get_worksheet(0)
            .err()
            .expect("internal hyperlink relationship must fail")
            .to_string();

        assert!(error.contains("not external"), "unexpected error: {error}");
    }
}
