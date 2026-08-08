//! OPC package assembly for XLSB workbooks.

use super::model::{SheetSlot, WorkbookWriter, XLSB_WORKSHEET_BINARY_INDEX_EMPTY};
use crate::package::error::{Error, Result};
use crate::package::formula::{
    CompilationContext, Context, DefinedName, ExternalSheet, SupportingLink,
};
use crate::raw::Writer;
use litchi_core::xml::escape_xml;
use litchi_opc::constants::{content_type as ct, relationship_type as rel};
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::io::{Seek, Write};
#[cfg(feature = "vba-inspection")]
use std::sync::Arc;

pub(super) fn checked_capacity(resource: &'static str, terms: &[usize]) -> Result<usize> {
    terms.iter().try_fold(0usize, |total, term| {
        total
            .checked_add(*term)
            .ok_or(Error::CapacityOverflow { resource })
    })
}

fn reserved_vec<T>(capacity: usize, resource: &'static str) -> Result<Vec<T>> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation { resource, source })?;
    Ok(values)
}

impl WorkbookWriter {
    /// Save the workbook to a writer
    ///
    /// # Arguments
    ///
    /// * `writer` - A writer that implements `Write` and `Seek`
    pub fn save<W: Write + Seek>(&mut self, writer: W) -> Result<()> {
        self.validate_formula_metadata()?;
        let mut xml_maps_plan = crate::writer::xml_maps::stage(
            self.xml_maps.as_ref(),
            self.connections.as_ref(),
            &self.worksheets,
        )?;
        let mut package = OpcPackage::new();

        // Add document properties (required by Excel)
        self.add_doc_props(&mut package)?;

        // Add theme (REQUIRED by Excel)
        self.add_theme(&mut package)?;

        // Add worksheets first so that shared_strings is fully populated before we
        // decide whether to create a sharedStrings part and relationship.
        let formula_sheet_ranges = self.add_worksheet_parts(&mut package, &mut xml_maps_plan)?;
        self.add_chart_sheet_parts(&mut package)?;

        // Add shared strings table only if non-empty. Excel-generated empty XLSB
        // workbooks omit sharedStrings.bin entirely, and the corresponding
        // relationship from the workbook.
        if !self.shared_strings.is_empty() {
            self.add_shared_strings_part(&mut package)?;
        }

        // Add styles
        self.add_styles_part(&mut package)?;

        // Finally add the workbook part (after worksheets / shared strings / styles)
        // so that relationships are created with full knowledge of which parts
        // actually exist.
        self.add_workbook_part(
            &mut package,
            &formula_sheet_ranges,
            xml_maps_plan.map_info_xml.is_some(),
        )?;

        if let Some(xml) = xml_maps_plan.map_info_xml.take() {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/xmlMaps.xml")?,
                litchi_ooxml_common::spreadsheet_xml_maps::CONTENT_TYPE.to_string(),
                xml,
            )));
        }

        for (index, link) in self.external_links.iter().enumerate() {
            let one_based_index = index.checked_add(1).ok_or_else(|| {
                Error::InvalidFormula("external-link part index overflow".to_string())
            })?;
            package.add_part(Box::new(crate::package::external_link::author_part(
                link,
                one_based_index,
            )?));
        }

        #[cfg(feature = "vba-inspection")]
        if let Some(payload) = &self.vba {
            crate::package::vba_project::store_vba_bytes(
                &mut package,
                &PackURI::new("/xl/workbook.bin")?,
                Arc::clone(payload),
            )?;
        }

        // External Data Connections part (at most one per package, related
        // from the workbook part).
        if let Some(connections) = &self.connections {
            let connections_uri = PackURI::new("/xl/connections.bin")?;
            package.add_part(Box::new(BlobPart::new(
                connections_uri.clone(),
                "application/vnd.ms-excel.connections".to_string(),
                crate::package::connections::write::write_connections_part(connections)?,
            )));
            package
                .get_part_mut(&PackURI::new("/xl/workbook.bin")?)?
                .rels_mut()
                .get_or_add(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections",
                    "connections.bin",
                );
        }

        // PivotCache Definition parts (one per attached cache, related from
        // the workbook part; the relationships and BrtBeginPivotCacheID
        // records were already emitted by add_workbook_part).
        for (index, cache) in self.pivot_caches.iter().enumerate() {
            let cache_uri = PackURI::new(format!(
                "/xl/pivotCache/pivotCacheDefinition{}.bin",
                index + 1
            ))?;
            package.add_part(Box::new(BlobPart::new(
                cache_uri,
                "application/vnd.ms-excel.pivotCacheDefinition".to_string(),
                cache.bytes.clone(),
            )));
        }

        // Save package to output
        package.to_stream(writer)?;

        Ok(())
    }

    fn authored_pivot_tables(&self) -> Vec<(String, String)> {
        self.worksheets
            .iter()
            .flat_map(|sheet| {
                sheet
                    .pivot_table_views()
                    .iter()
                    .map(move |view| (view.name().to_string(), sheet.name().to_string()))
            })
            .collect()
    }

    fn normalized_pivot_chart(
        chart: &crate::chart::Chart,
        host_sheet_name: &str,
        authored_pivot_tables: &[(String, String)],
    ) -> Result<crate::chart::Chart> {
        let Some(source) = chart.chart.pivot_source.as_ref() else {
            return Ok(chart.clone());
        };
        let name = crate::pivot_chart::resolve_authored_pivot_source_name(
            &source.name,
            host_sheet_name,
            authored_pivot_tables,
        )
        .map_err(|error| Error::InvalidFormula(error.to_string()))?;
        let mut normalized = chart.clone();
        normalized
            .chart
            .pivot_source
            .as_mut()
            .expect("pivot source presence checked above")
            .name = name;
        Ok(normalized)
    }

    /// Add document properties (required by Excel to open the file)
    fn add_doc_props(&self, package: &mut OpcPackage) -> Result<()> {
        // Add app.xml (Extended Properties)
        let app_xml = self.create_app_xml();
        let app_uri = PackURI::new("/docProps/app.xml")?;
        let app_part = BlobPart::new(
            app_uri,
            "application/vnd.openxmlformats-officedocument.extended-properties+xml".to_string(),
            app_xml.into_bytes(),
        );
        package.add_part(Box::new(app_part));
        package.relate_to(
            "docProps/app.xml",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        );

        // Add core.xml (Core Properties)
        let core_xml = self.create_core_xml();
        let core_uri = PackURI::new("/docProps/core.xml")?;
        let core_part = BlobPart::new(
            core_uri,
            "application/vnd.openxmlformats-package.core-properties+xml".to_string(),
            core_xml.as_bytes().to_vec(),
        );
        package.add_part(Box::new(core_part));
        package.relate_to(
            "docProps/core.xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        );

        Ok(())
    }

    /// Create app.xml content (Extended Properties)
    pub(super) fn create_app_xml(&self) -> String {
        let sheet_count = self.sheet_order.len();

        // Build sheet names list
        let mut sheet_names = String::new();
        for slot in &self.sheet_order {
            sheet_names.push_str(&format!(
                "<vt:lpstr>{}</vt:lpstr>",
                escape_xml(self.sheet_name(*slot))
            ));
        }

        crate::package::template::app(sheet_count, &sheet_names)
    }

    /// Create core.xml content (Core Properties)
    pub(super) fn create_core_xml(&self) -> &'static str {
        crate::package::template::core()
    }

    /// Add theme (REQUIRED by Excel to open file)
    fn add_theme(&self, package: &mut OpcPackage) -> Result<()> {
        // Create minimal Office theme
        let theme_xml = self.create_minimal_theme();
        let theme_uri = PackURI::new("/xl/theme/theme1.xml")?;
        let theme_part = BlobPart::new(
            theme_uri,
            "application/vnd.openxmlformats-officedocument.theme+xml".to_string(),
            theme_xml.as_bytes().to_vec(),
        );
        package.add_part(Box::new(theme_part));

        // Note: Relationship from workbook to theme will be added by workbook_part.rels_mut()
        Ok(())
    }

    /// Create minimal Office theme XML
    pub(super) fn create_minimal_theme(&self) -> &'static str {
        crate::package::template::theme()
    }

    /// Add workbook part to the package
    fn add_workbook_part(
        &self,
        package: &mut OpcPackage,
        formula_sheet_ranges: &[(u32, u32)],
        has_xml_maps: bool,
    ) -> Result<()> {
        // Create the workbook part with an empty blob first so that all
        // relationships are attached (with concrete IDs) before the workbook
        // stream is serialized: BrtBeginPivotCacheID records reference
        // relationship IDs.
        let workbook_uri = PackURI::new("/xl/workbook.bin")?;
        let mut workbook_part = BlobPart::new(
            workbook_uri.clone(),
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
            Vec::new(),
        );

        // Add relationships from workbook to worksheets and styles
        let mut pivot_cache_rel_ids = Vec::with_capacity(self.pivot_caches.len());
        let mut external_link_rel_ids = Vec::with_capacity(self.external_links.len());
        {
            let rels = workbook_part.rels_mut();
            for slot in &self.sheet_order {
                match *slot {
                    SheetSlot::Worksheet(index) => {
                        rels.get_or_add(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                            &format!("worksheets/sheet{}.bin", index + 1),
                        );
                    },
                    SheetSlot::ChartSheet(index) => {
                        rels.get_or_add(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
                            &format!("chartsheets/sheet{}.bin", index + 1),
                        );
                    },
                }
            }

            rels.get_or_add(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                "styles.bin",
            );

            // Add sharedStrings relationship only when the shared strings table is
            // non-empty. Excel omits sharedStrings.bin entirely for empty
            // workbooks, and the relationship MUST NOT reference a non-existent
            // part.
            if !self.shared_strings.is_empty() {
                rels.get_or_add(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                    "sharedStrings.bin",
                );
            }

            rels.get_or_add(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
                "theme/theme1.xml",
            );

            if has_xml_maps {
                rels.get_or_add(
                    litchi_ooxml_common::spreadsheet_xml_maps::REL,
                    "xmlMaps.xml",
                );
            }

            // PivotCache Definition relationships; the BrtBeginPivotCacheID
            // records below carry these relationship IDs.
            for (index, cache) in self.pivot_caches.iter().enumerate() {
                let relationship = rels.get_or_add(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
                    &format!("pivotCache/pivotCacheDefinition{}.bin", index + 1),
                );
                pivot_cache_rel_ids.push((cache.id, relationship.r_id().to_string()));
            }

            for (index, _) in self.external_links.iter().enumerate() {
                let relationship = rels.get_or_add(
                    rel::EXTERNAL_LINK,
                    &format!("externalLinks/externalLink{}.bin", index + 1),
                );
                external_link_rel_ids.push(relationship.r_id().to_string());
            }
        }

        // Write workbook structure
        let mut workbook_data = Vec::new();
        let mut writer = Writer::new(&mut workbook_data);
        self.write_workbook(
            &mut writer,
            formula_sheet_ranges,
            &pivot_cache_rel_ids,
            &external_link_rel_ids,
        )?;
        workbook_part.set_blob(workbook_data);

        // Add part to package
        package.add_part(Box::new(workbook_part));

        // Add relationship from root to workbook
        package.relate_to(
            "xl/workbook.bin",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        );

        Ok(())
    }

    fn add_worksheet_parts(
        &mut self,
        package: &mut OpcPackage,
        xml_maps_plan: &mut crate::writer::xml_maps::XmlMapsWritePlan,
    ) -> Result<Vec<(u32, u32)>> {
        let authored_pivot_tables = self.authored_pivot_tables();
        let worksheet_names = self
            .sheet_order
            .iter()
            .map(|slot| self.sheet_name(*slot).to_string())
            .collect::<Vec<_>>();
        let worksheet_sheet_indexes = self
            .sheet_order
            .iter()
            .enumerate()
            .filter_map(|(sheet_index, slot)| match slot {
                SheetSlot::Worksheet(worksheet_index) => Some((*worksheet_index, sheet_index)),
                SheetSlot::ChartSheet(_) => None,
            })
            .collect::<std::collections::HashMap<_, _>>();
        let defined_names = self
            .named_ranges
            .iter()
            .map(|named_range| DefinedName {
                name: named_range.name.clone(),
                sheet_id: named_range.sheet_id,
            })
            .collect::<Vec<_>>();
        let formula_sheet_ranges = std::cell::RefCell::new(Vec::new());
        // Compile every context-dependent formula before serializing any
        // worksheet. Besides making the final XTI table available to
        // sparkline validation, retaining the restore values keeps this
        // prepass transactional across every later package-writing failure.
        let mut compiled_formulas = reserved_vec(
            self.worksheets.len(),
            "sparkline formula preflight restorations",
        )?;
        for (i, worksheet) in self.worksheets.iter_mut().enumerate() {
            let current_sheet = u32::try_from(worksheet_sheet_indexes[&i])
                .map_err(|_| Error::InvalidFormula("worksheet index overflow".to_string()))?;
            let formula_context = CompilationContext {
                worksheet_names: &worksheet_names,
                defined_names: &defined_names,
                tables: &[],
                supporting_links: &[],
                external_sheets: &[],
                external_books: &[],
                sheet_ranges: &formula_sheet_ranges,
                current_sheet,
            };
            match worksheet.compile_contextual_formulas(&formula_context) {
                Ok(restore) => compiled_formulas.push(restore),
                Err(error) => {
                    for (worksheet, restore) in
                        self.worksheets.iter_mut().zip(compiled_formulas.drain(..))
                    {
                        worksheet.clear_compiled_formulas(restore);
                    }
                    return Err(error);
                },
            }
        }

        let mut next_table_index = 1usize;
        let mut next_drawing_index = 1usize;
        let mut next_chart_index = 1usize;
        let mut next_image_index = 1usize;
        let mut next_pivot_table_index = 1usize;
        let mut next_single_cell_index = 1usize;
        let write_result = (|| -> Result<()> {
            let mut supporting_links = reserved_vec(1, "sparkline supporting-link context")?;
            supporting_links.push(SupportingLink::SelfWorkbook);

            let ranges = formula_sheet_ranges.borrow();
            let external_sheet_capacity = checked_capacity(
                "sparkline XTI context",
                &[2, self.sheet_order.len(), ranges.len()],
            )?;
            let mut external_sheets =
                reserved_vec(external_sheet_capacity, "sparkline XTI context")?;
            external_sheets.push(ExternalSheet {
                external_link: 0,
                first_sheet: -2,
                last_sheet: -2,
            });
            external_sheets.push(ExternalSheet {
                external_link: 0,
                first_sheet: -1,
                last_sheet: -1,
            });
            for sheet in 0..self.sheet_order.len() {
                let sheet = i32::try_from(sheet)
                    .map_err(|_| Error::InvalidFormula("worksheet index overflow".to_string()))?;
                external_sheets.push(ExternalSheet {
                    external_link: 0,
                    first_sheet: sheet,
                    last_sheet: sheet,
                });
            }
            for &(first, last) in ranges.iter() {
                external_sheets.push(ExternalSheet {
                    external_link: 0,
                    first_sheet: i32::try_from(first).map_err(|_| {
                        Error::InvalidFormula("first formula sheet index overflow".to_string())
                    })?,
                    last_sheet: i32::try_from(last).map_err(|_| {
                        Error::InvalidFormula("last formula sheet index overflow".to_string())
                    })?,
                });
            }
            drop(ranges);

            // Sparkline context validation only inspects these collection
            // lengths. Avoid cloning user-controlled names solely to prove
            // one-based indexes.
            let mut context_worksheet_names =
                reserved_vec(worksheet_names.len(), "sparkline worksheet-name context")?;
            context_worksheet_names.resize_with(worksheet_names.len(), String::new);
            let mut context_defined_names =
                reserved_vec(self.named_ranges.len(), "sparkline defined-name context")?;
            context_defined_names.resize_with(self.named_ranges.len(), String::new);

            let sparkline_context = Context {
                worksheet_names: context_worksheet_names.into(),
                supporting_links: supporting_links.into(),
                external_sheets: external_sheets.into(),
                external_books: Vec::new().into(),
                defined_names: context_defined_names.into(),
                tables: Vec::new().into(),
                pivot_views: Vec::new().into(),
                pivot_name_scopes: Vec::new().into(),
                active_pivot_scope: None,
                current_sheet: None,
            };
            for worksheet in &self.worksheets {
                let groups = worksheet.sparkline_groups();
                if groups.is_some_and(|groups| {
                    groups.iter().any(|group| {
                        group.date_formula().is_some_and(|formula| {
                            formula.kind() == crate::sparkline::FormulaKind::ExternalName
                        }) || group.sparklines().iter().any(|sparkline| {
                            sparkline.formula().is_some_and(|formula| {
                                formula.kind() == crate::sparkline::FormulaKind::ExternalName
                            })
                        })
                    })
                }) {
                    return Err(Error::UnsupportedFeature(
                        "new XLSB writer cannot author the external-workbook BrtExternSheet entry required by sparkline PtgNameX"
                            .to_string(),
                    ));
                }
                crate::sparkline::workbook::validate_groups_context(groups, &sparkline_context)?;
            }

            // Stage every block before the first per-sheet relationship ID,
            // shared string, table ID, or drawing ID can be published.
            let mut sparkline_blocks =
                reserved_vec(self.worksheets.len(), "staged worksheet sparkline blocks")?;
            for worksheet in &self.worksheets {
                sparkline_blocks.push(worksheet.stage_sparkline_block()?);
            }

            for (i, worksheet) in self.worksheets.iter_mut().enumerate() {
                // Create the worksheet part with an empty blob first so we can attach
                // relationships (binary index + external hyperlinks) and obtain
                // concrete relationship IDs before serializing the sheet data.
                let sheet_uri = PackURI::new(format!("/xl/worksheets/sheet{}.bin", i + 1))?;
                let mut sheet_part = BlobPart::new(
                    sheet_uri,
                    "application/vnd.ms-excel.worksheet".to_string(),
                    Vec::new(),
                );

                // Each worksheet MUST have a Worksheet Binary Index part. Excel adds
                // this automatically when repairing our files. We proactively create
                // it here and wire up the relationship so the package is valid
                // without requiring Excel repair.
                let binary_index_name = format!("binaryIndex{}.bin", i + 1);
                let binary_index_uri =
                    PackURI::new(format!("/xl/worksheets/{}", binary_index_name))?;
                let binary_index_part = BlobPart::new(
                    binary_index_uri,
                    "application/vnd.ms-excel.binIndexWs".to_string(),
                    XLSB_WORKSHEET_BINARY_INDEX_EMPTY.to_vec(),
                );

                {
                    let rels = sheet_part.rels_mut();
                    rels.get_or_add(
                        "http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex",
                        &binary_index_name,
                    );
                }

                // Create external hyperlink relationships and record their rIds
                // back into the worksheet's Hyperlink structs so that the
                // subsequent BrtHLink records carry valid relationship IDs.
                for hyperlink in worksheet.hyperlinks_mut() {
                    if let Some(ref target) = hyperlink.target
                        && (target.starts_with("http://")
                            || target.starts_with("https://")
                            || target.starts_with("ftp://")
                            || target.starts_with("mailto:"))
                    {
                        let rel_id = sheet_part.relate_to_ext(target, rel::HYPERLINK);
                        hyperlink.r_id = rel_id;
                    }
                }

                let comments_part = if worksheet.comments().is_empty() {
                    None
                } else {
                    let comments_name = format!("comments{}.bin", i + 1);
                    sheet_part.relate_to(&format!("../{comments_name}"), rel::COMMENTS);
                    let mut comments_data = Vec::new();
                    crate::comments::write(
                        &mut Writer::new(&mut comments_data),
                        worksheet.comments(),
                    )?;
                    Some(BlobPart::new(
                        PackURI::new(format!("/xl/{comments_name}"))?,
                        "application/vnd.ms-excel.comments".to_string(),
                        comments_data,
                    ))
                };

                // Structured tables: one part per table, related from the
                // worksheet so the BrtListPart records can carry valid rIds.
                let mut table_parts = Vec::new();
                if !worksheet.tables().is_empty() {
                    let mut rel_ids = Vec::with_capacity(worksheet.tables().len());
                    for table_ordinal in 0..worksheet.tables().len() {
                        let table_name = format!("tables/table{next_table_index}.bin");
                        next_table_index += 1;
                        rel_ids.push(sheet_part.relate_to(&format!("../{table_name}"), rel::TABLE));
                        table_parts.push(BlobPart::new(
                            PackURI::new(format!("/xl/{table_name}"))?,
                            "application/vnd.ms-excel.table".to_string(),
                            std::mem::take(
                                &mut xml_maps_plan.worksheets[i].table_parts[table_ordinal],
                            ),
                        ));
                    }
                    worksheet.table_rel_ids = rel_ids;
                }

                let single_cell_part = if let Some(bytes) =
                    xml_maps_plan.worksheets[i].single_cells.take()
                {
                    let name = format!("tableSingleCells{next_single_cell_index}.bin");
                    next_single_cell_index =
                        next_single_cell_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormula("single-cell XML part index overflow".to_string())
                        })?;
                    sheet_part.relate_to(
                            &format!("../tables/{name}"),
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells",
                        );
                    Some(BlobPart::new(
                        PackURI::new(format!("/xl/tables/{name}"))?,
                        "application/vnd.ms-excel.tableSingleCells".to_string(),
                        bytes,
                    ))
                } else {
                    None
                };

                // PivotTable definitions are related implicitly from their host
                // worksheet and back to the exact workbook PivotCache definition.
                let mut pivot_table_parts = Vec::new();
                for view in worksheet.pivot_table_views() {
                    let cache_index = self
                        .pivot_caches
                        .iter()
                        .position(|cache| cache.id == view.cache_id())
                        .ok_or_else(|| {
                            Error::InvalidFormula(format!(
                                "PivotTable view {:?} references unknown cache {}",
                                view.name(),
                                view.cache_id()
                            ))
                        })?;
                    let cache_version_created = self.pivot_caches[cache_index].version_created;
                    if (view.version_created() >= 3) != (cache_version_created >= 3) {
                        return Err(Error::InvalidFormula(format!(
                            "PivotTable view {:?} functionality level {} is incompatible with cache {} level {}",
                            view.name(),
                            view.version_created(),
                            view.cache_id(),
                            cache_version_created
                        )));
                    }
                    let pivot_name = format!("pivotTable{next_pivot_table_index}.bin");
                    next_pivot_table_index =
                        next_pivot_table_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormula("PivotTable part index overflow".to_string())
                        })?;
                    sheet_part.relate_to(&format!("../pivotTables/{pivot_name}"), rel::PIVOT_TABLE);
                    let mut part = BlobPart::new(
                        PackURI::new(format!("/xl/pivotTables/{pivot_name}"))?,
                        "application/vnd.ms-excel.PivotTable".to_string(),
                        view.as_bytes().to_vec(),
                    );
                    part.relate_to(
                        &format!("../pivotCache/pivotCacheDefinition{}.bin", cache_index + 1),
                        rel::PIVOT_CACHE_DEFINITION,
                    );
                    pivot_table_parts.push(part);
                }

                // Worksheet images and charts use standard SpreadsheetDrawing,
                // Image, and DrawingML Chart parts. The binary sheet carries only
                // BrtDrawing with the relationship ID allocated here.
                let mut drawing_part = None;
                let mut chart_parts = Vec::new();
                let mut image_parts = Vec::new();
                if worksheet.has_drawing_objects() {
                    let normalized_charts = worksheet
                        .charts()
                        .iter()
                        .map(|chart| {
                            Self::normalized_pivot_chart(
                                chart,
                                worksheet.name(),
                                &authored_pivot_tables,
                            )
                        })
                        .collect::<Result<Vec<_>>>()?;
                    let drawing_xml = crate::package::drawing_write::serialize_drawing(
                        worksheet.images(),
                        &normalized_charts,
                        worksheet.shapes(),
                        worksheet.groups(),
                        worksheet.connections(),
                    )?;
                    let drawing_name = format!("drawing{next_drawing_index}.xml");
                    next_drawing_index = next_drawing_index.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormula("drawing part index overflow".to_string())
                    })?;
                    let mut part = BlobPart::new(
                        PackURI::new(format!("/xl/drawings/{drawing_name}"))?,
                        ct::OFC_DRAWING.to_string(),
                        drawing_xml,
                    );
                    for (image_ordinal, image) in worksheet.images().iter().enumerate() {
                        let image_name =
                            format!("image{next_image_index}.{}", image.format().extension());
                        next_image_index = next_image_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormula("image part index overflow".to_string())
                        })?;
                        let relationship_id =
                            part.relate_to(&format!("../media/{image_name}"), rel::IMAGE);
                        let expected_relationship_id = format!("rId{}", image_ordinal + 1);
                        if relationship_id != expected_relationship_id {
                            return Err(Error::InvalidFormula(format!(
                                "drawing image relationship allocation mismatch: expected {expected_relationship_id}, got {relationship_id}"
                            )));
                        }
                        image_parts.push(BlobPart::new(
                            PackURI::new(format!("/xl/media/{image_name}"))?,
                            image.format().content_type().to_string(),
                            image.data().to_vec(),
                        ));
                    }
                    for (chart_ordinal, chart) in normalized_charts.iter().enumerate() {
                        let chart_name = format!("chart{next_chart_index}.xml");
                        let graph = crate::package::chart_resources::author_chart_graph(
                            chart,
                            next_chart_index,
                        )?;
                        next_chart_index = next_chart_index.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormula("chart part index overflow".to_string())
                        })?;
                        let relationship_id =
                            part.relate_to(&format!("../charts/{chart_name}"), rel::CHART);
                        let expected_relationship_id =
                            format!("rId{}", worksheet.images().len() + chart_ordinal + 1);
                        if relationship_id != expected_relationship_id {
                            return Err(Error::InvalidFormula(format!(
                                "drawing chart relationship allocation mismatch: expected {expected_relationship_id}, got {relationship_id}"
                            )));
                        }
                        chart_parts.push(graph.chart_part);
                        chart_parts.extend(graph.related_parts);
                    }
                    let rel_id =
                        sheet_part.relate_to(&format!("../drawings/{drawing_name}"), rel::DRAWING);
                    worksheet.set_drawing_rel_id(Some(rel_id));
                    drawing_part = Some(part);
                } else {
                    worksheet.set_drawing_rel_id(None);
                }

                // Now serialize the worksheet with fully-populated relationship IDs
                // in the hyperlink records.
                let mut sheet_data = Vec::new();
                let worksheet_write_result = {
                    let mut writer = Writer::new(&mut sheet_data);
                    worksheet.write_with_sparkline_block(
                        &mut writer,
                        &mut self.shared_strings,
                        sparkline_blocks[i].as_deref(),
                    )
                };
                worksheet_write_result?;
                sheet_part.set_blob(sheet_data);

                package.add_part(Box::new(sheet_part));
                package.add_part(Box::new(binary_index_part));
                if let Some(part) = comments_part {
                    package.add_part(Box::new(part));
                }
                for part in table_parts {
                    package.add_part(Box::new(part));
                }
                if let Some(part) = single_cell_part {
                    package.add_part(Box::new(part));
                }
                for part in pivot_table_parts {
                    package.add_part(Box::new(part));
                }
                if let Some(part) = drawing_part {
                    package.add_part(Box::new(part));
                }
                for part in chart_parts {
                    package.add_part(Box::new(part));
                }
                for part in image_parts {
                    package.add_part(Box::new(part));
                }
            }
            Ok(())
        })();

        for (worksheet, restore) in self
            .worksheets
            .iter_mut()
            .zip(compiled_formulas.into_iter())
        {
            worksheet.clear_compiled_formulas(restore);
        }
        write_result?;

        Ok(formula_sheet_ranges.into_inner())
    }

    /// Add binary Chart Sheet streams and their standard DrawingML graphs.
    fn add_chart_sheet_parts(&self, package: &mut OpcPackage) -> Result<()> {
        let mut next_drawing_index = self
            .worksheets
            .iter()
            .filter(|sheet| sheet.has_drawing_objects())
            .count()
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormula("drawing part index overflow".to_string()))?;
        let mut next_chart_index = self
            .worksheets
            .iter()
            .try_fold(1usize, |next, sheet| next.checked_add(sheet.charts().len()))
            .ok_or_else(|| Error::InvalidFormula("chart part index overflow".to_string()))?;

        for (index, sheet) in self.chart_sheets.iter().enumerate() {
            sheet.validate()?;
            let drawing_name = format!("drawing{next_drawing_index}.xml");
            next_drawing_index = next_drawing_index
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormula("drawing part index overflow".to_string()))?;
            let chart_index = next_chart_index;
            let chart_name = format!("chart{chart_index}.xml");
            next_chart_index = next_chart_index
                .checked_add(1)
                .ok_or_else(|| Error::InvalidFormula("chart part index overflow".to_string()))?;
            let normalized_chart = Self::normalized_pivot_chart(
                sheet.chart(),
                sheet.name(),
                &self.authored_pivot_tables(),
            )?;
            let graph = crate::package::chart_resources::author_chart_graph(
                &normalized_chart,
                chart_index,
            )?;

            let mut chart_sheet_part = BlobPart::new(
                PackURI::new(format!("/xl/chartsheets/sheet{}.bin", index + 1))?,
                "application/vnd.ms-excel.chartsheet".to_string(),
                Vec::new(),
            );
            let drawing_rel_id =
                chart_sheet_part.relate_to(&format!("../drawings/{drawing_name}"), rel::DRAWING);
            if drawing_rel_id != "rId1" {
                return Err(Error::InvalidFormula(format!(
                    "chart-sheet drawing relationship allocation mismatch: {drawing_rel_id}"
                )));
            }

            let mut printer_part = None;
            let printer_rel_id = if let Some(bytes) = sheet.printer_settings() {
                let printer_name = format!("printerSettings{}.bin", index + 1);
                let rel_id = chart_sheet_part.relate_to(
                    &format!("../printerSettings/{printer_name}"),
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings",
                );
                printer_part = Some(BlobPart::new(
                    PackURI::new(format!("/xl/printerSettings/{printer_name}"))?,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"
                        .to_string(),
                    bytes.to_vec(),
                ));
                Some(rel_id)
            } else {
                None
            };
            chart_sheet_part.set_blob(crate::writer::chartsheet::write_chart_sheet(
                sheet,
                &drawing_rel_id,
                printer_rel_id.as_deref(),
            )?);

            let mut drawing_part = BlobPart::new(
                PackURI::new(format!("/xl/drawings/{drawing_name}"))?,
                ct::OFC_DRAWING.to_string(),
                crate::package::drawing_write::serialize_chart_sheet_drawing(sheet.name())?,
            );
            let chart_rel_id =
                drawing_part.relate_to(&format!("../charts/{chart_name}"), rel::CHART);
            if chart_rel_id != "rId1" {
                return Err(Error::InvalidFormula(format!(
                    "chart-sheet chart relationship allocation mismatch: {chart_rel_id}"
                )));
            }
            package.add_part(Box::new(chart_sheet_part));
            package.add_part(Box::new(drawing_part));
            package.add_part(Box::new(graph.chart_part));
            for part in graph.related_parts {
                package.add_part(Box::new(part));
            }
            if let Some(part) = printer_part {
                package.add_part(Box::new(part));
            }
        }
        Ok(())
    }

    /// Add shared strings part to the package
    fn add_shared_strings_part(&self, package: &mut OpcPackage) -> Result<()> {
        let mut sst_data = Vec::new();
        let mut writer = Writer::new(&mut sst_data);

        self.shared_strings.write(&mut writer)?;

        let sst_uri = PackURI::new("/xl/sharedStrings.bin")?;
        let sst_part = BlobPart::new(
            sst_uri,
            "application/vnd.ms-excel.sharedStrings".to_string(),
            sst_data,
        );

        package.add_part(Box::new(sst_part));

        Ok(())
    }

    /// Add styles part to the package
    fn add_styles_part(&self, package: &mut OpcPackage) -> Result<()> {
        let mut styles_data = Vec::new();
        let mut writer = Writer::new(&mut styles_data);

        self.styles.write(&mut writer)?;

        let styles_uri = PackURI::new("/xl/styles.bin")?;
        let styles_part = BlobPart::new(
            styles_uri,
            "application/vnd.ms-excel.styles".to_string(),
            styles_data,
        );

        package.add_part(Box::new(styles_part));

        Ok(())
    }
}
