//! XLSB workbook writer implementation
//!
//! This module provides functionality to create complete XLSB files with multiple worksheets,
//! shared strings, styles, and advanced features.
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{
    CellParsedFormula, FormulaCompilationContext, FormulaDefinedName, excel_name_eq,
};
use crate::xlsb::named_ranges::{NamedRange, validate_defined_name};
use crate::xlsb::writer::{
    MutableSharedStringsWriter, MutableXlsbChartSheet, MutableXlsbWorksheet, StylesWriter,
};
use litchi_core::xml::escape_xml;
use litchi_opc::constants::{content_type as ct, relationship_type as rel};
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use litchi_xlsb::calc::{self, Props};
use litchi_xlsb::raw::{Writer, kind};
use std::io::{Seek, Write};
use std::sync::Arc;

const MAX_AUTHORED_EXTERNAL_LINKS: usize = 4_096;

/// XLSB workbook writer
///
/// Creates complete XLSB workbook files with support for:
/// - Multiple worksheets
/// - Shared strings
/// - Styles (fonts, fills, borders, number formats)
/// - Workbook properties (date system, etc.)
///
/// # Example
///
/// ```rust,no_run
/// use litchi_ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
/// use std::fs::File;
///
/// let mut workbook = XlsbWorkbookWriter::new();
///
/// let mut sheet = MutableXlsbWorksheet::new("Sheet1");
/// sheet.set_cell(0, 0, "Hello");
/// sheet.set_cell(0, 1, 42.0);
///
/// workbook.add_worksheet(sheet);
///
/// let file = File::create("output.xlsb")?;
/// workbook.save(file)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct XlsbWorkbookWriter {
    worksheets: Vec<MutableXlsbWorksheet>,
    chart_sheets: Vec<MutableXlsbChartSheet>,
    sheet_order: Vec<XlsbSheetSlot>,
    named_ranges: Vec<NamedRange>,
    shared_strings: MutableSharedStringsWriter,
    styles: StylesWriter,
    calc: Props,
    is_1904: bool,
    connections: Option<crate::xlsb::connections::XlsbConnections>,
    external_links: Vec<crate::xlsb::formula::XlsbExternalLink>,
    pivot_caches: Vec<AuthoredPivotCache>,
    vba: Option<Arc<Vec<u8>>>,
}

#[derive(Debug, Clone, Copy)]
enum XlsbSheetSlot {
    Worksheet(usize),
    ChartSheet(usize),
}

struct AuthoredPivotCache {
    id: u32,
    version_created: u8,
    bytes: Vec<u8>,
}

/// Minimal Worksheet Binary Index payload for an empty worksheet.
///
/// This binary blob was captured from an Excel-generated empty XLSB file
/// (`excel_empty.xlsb`) and represents a valid Worksheet Binary Index part
/// for a simple sheet without additional features. According to
/// [MS-XLSB] 2.1.7.63 (Worksheet Binary Index), a worksheet MUST have a
/// corresponding binary index part.
///
/// TODO: If we start emitting advanced worksheet features that rely on the
/// binary index (for example, very large sheets or complex structures),
/// this payload should be generated from the official ABNF grammar instead
/// of using this minimal fixed template.
const XLSB_WORKSHEET_BINARY_INDEX_EMPTY: [u8; 29] = [
    0x2a, 0x18, 0x00, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x95, 0x02, 0x00,
];

impl XlsbWorkbookWriter {
    /// Create a new XLSB workbook writer
    pub fn new() -> Self {
        XlsbWorkbookWriter {
            worksheets: Vec::new(),
            chart_sheets: Vec::new(),
            sheet_order: Vec::new(),
            named_ranges: Vec::new(),
            shared_strings: MutableSharedStringsWriter::new(),
            styles: StylesWriter::new(),
            calc: Props::default(),
            is_1904: false,
            connections: None,
            external_links: Vec::new(),
            pivot_caches: Vec::new(),
            vba: None,
        }
    }

    /// Set the date system (1900 or 1904)
    ///
    /// # Arguments
    ///
    /// * `is_1904` - `true` for 1904 date system (Mac), `false` for 1900 (Windows, default)
    pub fn set_date_system(&mut self, is_1904: bool) {
        self.is_1904 = is_1904;
    }

    /// Validated workbook formula calculation policy written to `BrtCalcProp`.
    pub fn calc(&self) -> &Props {
        &self.calc
    }

    /// Mutably configure the policy through its safe setters.
    pub fn calc_mut(&mut self) -> &mut Props {
        &mut self.calc
    }

    /// Replace the policy by moving in an already validated value.
    pub fn put_calc(&mut self, props: Props) -> &mut Self {
        self.calc = props;
        self
    }

    /// Attach a cache-free, inert MS-OVBA project to generated workbooks.
    pub fn set_vba(&mut self, project: litchi_vba::build::Project) -> XlsbResult<&mut Self> {
        self.set_vba_with(project, &litchi_vba::Limits::default())
    }

    /// Attach a cache-free project with explicit resource limits.
    pub fn set_vba_with(
        &mut self,
        project: litchi_vba::build::Project,
        limits: &litchi_vba::Limits,
    ) -> XlsbResult<&mut Self> {
        let payload = project.finish(limits)?;
        Ok(self.put_vba(payload))
    }

    /// Attach a prevalidated `vbaProject.bin` payload without executing it.
    pub fn put_vba(&mut self, payload: litchi_vba::Payload) -> &mut Self {
        self.vba = Some(Arc::new(payload.into_bytes()));
        self
    }

    /// Remove the project scheduled for insertion into generated workbooks.
    pub fn clear_vba(&mut self) -> bool {
        self.vba.take().is_some()
    }

    /// Add a worksheet to the workbook
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
    ///
    /// let mut workbook = XlsbWorkbookWriter::new();
    /// let sheet = MutableXlsbWorksheet::new("Sheet1");
    /// workbook.add_worksheet(sheet);
    /// ```
    pub fn add_worksheet(&mut self, worksheet: MutableXlsbWorksheet) {
        self.worksheets.push(worksheet);
        self.sheet_order
            .push(XlsbSheetSlot::Worksheet(self.worksheets.len() - 1));
    }

    /// Add a chart sheet in the current workbook sheet order.
    ///
    /// The chart is stored inertly in standard DrawingML parts. Package
    /// relationships are allocated at save time and never followed or fetched.
    pub fn add_chart_sheet(
        &mut self,
        chart_sheet: MutableXlsbChartSheet,
    ) -> XlsbResult<&mut MutableXlsbChartSheet> {
        chart_sheet.validate()?;
        if self.chart_sheets.len() >= super::chartsheet::max_chart_sheets() {
            return Err(XlsbError::InvalidFormula(
                "XLSB chart-sheet count limit exceeded".to_string(),
            ));
        }
        if self
            .sheet_order
            .iter()
            .any(|slot| excel_name_eq(self.sheet_name(*slot), chart_sheet.name()))
        {
            return Err(XlsbError::InvalidFormula(format!(
                "duplicate sheet name {:?}",
                chart_sheet.name()
            )));
        }
        self.chart_sheets.push(chart_sheet);
        self.sheet_order
            .push(XlsbSheetSlot::ChartSheet(self.chart_sheets.len() - 1));
        Ok(self
            .chart_sheets
            .last_mut()
            .expect("chart sheet was just added"))
    }

    /// Add a named range (defined name) to the workbook.
    pub fn add_named_range(&mut self, named_range: NamedRange) {
        self.named_ranges.push(named_range);
    }

    /// Attach an External Data Connections part (MS-XLSB 2.1.7.24) to the
    /// workbook.
    ///
    /// Connection identifiers and names must be unique and non-empty;
    /// strings, commands, URLs, and credential metadata are stored verbatim
    /// and are never resolved, contacted, refreshed, or executed.
    pub fn set_connections(
        &mut self,
        connections: crate::xlsb::connections::XlsbConnections,
    ) -> XlsbResult<()> {
        crate::xlsb::connections::write::validate_connections(&connections)?;
        self.connections = Some(connections);
        Ok(())
    }

    /// The attached External Data Connections part, when set.
    pub fn connections(&self) -> Option<&crate::xlsb::connections::XlsbConnections> {
        self.connections.as_ref()
    }

    /// Add inert Workbook, DDE, or OLE external-link metadata.
    ///
    /// External targets are stored but never followed, contacted, refreshed,
    /// instantiated, evaluated, or executed.
    pub fn add_external_link(
        &mut self,
        link: crate::xlsb::formula::XlsbExternalLink,
    ) -> XlsbResult<&mut Self> {
        link.validate()?;
        if self.external_links.len() >= MAX_AUTHORED_EXTERNAL_LINKS {
            return Err(XlsbError::InvalidFormula(
                "XLSB external-link count exceeds the safety limit".to_string(),
            ));
        }
        self.external_links.push(link);
        Ok(self)
    }

    /// External links scheduled for authoring, in workbook support-link order.
    pub fn external_links(&self) -> &[crate::xlsb::formula::XlsbExternalLink] {
        &self.external_links
    }

    /// Attach a PivotCache definition (MS-XLSB 2.1.7.38) to the workbook.
    ///
    /// The definition is serialized immediately, so model content the
    /// serializer cannot represent losslessly is rejected here rather than
    /// at save time. Returns the allocated workbook PivotCache identifier
    /// (`idSx`), unique per writer. Cache contents are stored verbatim and
    /// are never refreshed, contacted, or evaluated.
    pub fn add_pivot_cache(
        &mut self,
        definition: &crate::xlsb::pivot::PivotCacheDefinition,
    ) -> XlsbResult<u32> {
        let bytes = crate::xlsb::pivot::write::write_pivot_cache_definition(definition)?;
        let cache_id = u32::try_from(self.pivot_caches.len())
            .ok()
            .and_then(|next| next.checked_add(1))
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "PivotCache identifier overflow".to_string(),
                )
            })?;
        self.pivot_caches.push(AuthoredPivotCache {
            id: cache_id,
            version_created: definition.version_created,
            bytes,
        });
        Ok(cache_id)
    }

    /// Get a mutable reference to a worksheet by index
    pub fn get_worksheet_mut(&mut self, index: usize) -> Option<&mut MutableXlsbWorksheet> {
        self.worksheets.get_mut(index)
    }

    /// Get the number of worksheets
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Number of chart sheets in the workbook.
    pub fn chart_sheet_count(&self) -> usize {
        self.chart_sheets.len()
    }

    /// Mutable chart-sheet access by chart-sheet insertion index.
    pub fn get_chart_sheet_mut(&mut self, index: usize) -> Option<&mut MutableXlsbChartSheet> {
        self.chart_sheets.get_mut(index)
    }

    fn sheet_name(&self, slot: XlsbSheetSlot) -> &str {
        match slot {
            XlsbSheetSlot::Worksheet(index) => self.worksheets[index].name(),
            XlsbSheetSlot::ChartSheet(index) => self.chart_sheets[index].name(),
        }
    }

    /// Get a reference to the styles writer
    pub fn styles(&self) -> &StylesWriter {
        &self.styles
    }

    /// Get a mutable reference to the styles writer
    pub fn styles_mut(&mut self) -> &mut StylesWriter {
        &mut self.styles
    }

    /// Save the workbook to a writer
    ///
    /// # Arguments
    ///
    /// * `writer` - A writer that implements `Write` and `Seek`
    pub fn save<W: Write + Seek>(&mut self, writer: W) -> XlsbResult<()> {
        self.validate_formula_metadata()?;
        let mut package = OpcPackage::new();

        // Add document properties (required by Excel)
        self.add_doc_props(&mut package)?;

        // Add theme (REQUIRED by Excel)
        self.add_theme(&mut package)?;

        // Add worksheets first so that shared_strings is fully populated before we
        // decide whether to create a sharedStrings part and relationship.
        let formula_sheet_ranges = self.add_worksheet_parts(&mut package)?;
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
        self.add_workbook_part(&mut package, &formula_sheet_ranges)?;

        for (index, link) in self.external_links.iter().enumerate() {
            let one_based_index = index.checked_add(1).ok_or_else(|| {
                XlsbError::InvalidFormula("external-link part index overflow".to_string())
            })?;
            package.add_part(Box::new(
                crate::xlsb::external_link_write::author_external_link_part(link, one_based_index)?,
            ));
        }

        if let Some(payload) = &self.vba {
            crate::xlsb::vba_project::store_vba_bytes(
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
                crate::xlsb::connections::write::write_connections_part(connections)?,
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

    fn validate_formula_metadata(&self) -> XlsbResult<()> {
        if self.sheet_order.len() > usize::from(u16::MAX) - 2 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "{} sheets exceed the XLSB extern-sheet index limit",
                self.sheet_order.len()
            )));
        }
        for (index, slot) in self.sheet_order.iter().copied().enumerate() {
            let name = self.sheet_name(slot);
            let name_len = name.encode_utf16().count();
            if name_len == 0
                || name_len > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "sheet name {name:?} does not follow BrtBundleSh grammar"
                )));
            }
            if self.sheet_order[..index]
                .iter()
                .any(|existing| excel_name_eq(self.sheet_name(*existing), name))
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "duplicate sheet name {name:?}"
                )));
            }
        }
        for chart_sheet in &self.chart_sheets {
            chart_sheet.validate()?;
        }
        for (index, named_range) in self.named_ranges.iter().enumerate() {
            if named_range.function {
                return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                    "macro defined name {} cannot be emitted",
                    named_range.name
                )));
            }
            validate_defined_name(&named_range.name)?;
            if named_range.formula.is_none() {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "defined name {} has no formula",
                    named_range.name
                )));
            }
            if named_range.sheet_id.is_some_and(|sheet_id| {
                usize::try_from(sheet_id)
                    .ok()
                    .is_none_or(|sheet_id| sheet_id >= self.sheet_order.len())
            }) {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "defined name {} has invalid sheet scope {:?}",
                    named_range.name, named_range.sheet_id
                )));
            }
            if self.named_ranges[..index].iter().any(|existing| {
                existing.sheet_id == named_range.sheet_id
                    && excel_name_eq(&existing.name, &named_range.name)
            }) {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "duplicate defined name {:?} in scope {:?}",
                    named_range.name, named_range.sheet_id
                )));
            }
        }
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
        chart: &crate::xlsx::WorksheetChart,
        host_sheet_name: &str,
        authored_pivot_tables: &[(String, String)],
    ) -> XlsbResult<crate::xlsx::WorksheetChart> {
        let Some(source) = chart.chart.pivot_source.as_ref() else {
            return Ok(chart.clone());
        };
        let name = crate::xlsx::pivot_chart::resolve_authored_pivot_source_name(
            &source.name,
            host_sheet_name,
            authored_pivot_tables,
        )
        .map_err(|error| XlsbError::InvalidFormula(error.to_string()))?;
        let mut normalized = chart.clone();
        normalized
            .chart
            .pivot_source
            .as_mut()
            .expect("pivot source presence checked above")
            .name = name;
        Ok(normalized)
    }

    /// Write workbook-level defined names (BrtName records).
    fn write_named_ranges<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        for named_range in &self.named_ranges {
            if named_range.function {
                return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                    "macro defined name {} cannot be emitted",
                    named_range.name
                )));
            }
            validate_defined_name(&named_range.name)?;
            if let Some(sheet_id) = named_range.sheet_id
                && usize::try_from(sheet_id)
                    .ok()
                    .is_none_or(|index| index >= self.sheet_order.len())
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "defined name {} has invalid sheet scope {sheet_id}",
                    named_range.name
                )));
            }
            let formula = named_range.formula.as_ref().ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "defined name {} has no formula",
                    named_range.name
                ))
            })?;
            let parsed_formula = CellParsedFormula {
                rgce: formula.clone(),
                rgcb: Vec::new(),
            };
            let mut data = Vec::new();
            let mut temp_writer = Writer::new(&mut data);

            let mut flags = 0u32;
            if named_range.hidden {
                flags |= 0x0001;
            }
            temp_writer.write_u32(flags)?;
            temp_writer.write_u8(0)?; // chKey; zero for non-macro names

            temp_writer.write_u32(named_range.sheet_id.unwrap_or(u32::MAX))?;
            temp_writer.write_wide_string(&named_range.name)?;
            for byte in parsed_formula.to_bytes()? {
                temp_writer.write_u8(byte)?;
            }
            temp_writer.write_u32(u32::MAX)?; // NULL comment

            writer.write_record(kind::NAME, &data)?;
        }

        Ok(())
    }

    // Content types are handled automatically by the OPC package

    /// Add document properties (required by Excel to open the file)
    fn add_doc_props(&self, package: &mut OpcPackage) -> XlsbResult<()> {
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
            core_xml.into_bytes(),
        );
        package.add_part(Box::new(core_part));
        package.relate_to(
            "docProps/core.xml",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        );

        Ok(())
    }

    /// Create app.xml content (Extended Properties)
    fn create_app_xml(&self) -> String {
        let sheet_count = self.sheet_order.len();

        // Build sheet names list
        let mut sheet_names = String::new();
        for slot in &self.sheet_order {
            sheet_names.push_str(&format!(
                "<vt:lpstr>{}</vt:lpstr>",
                escape_xml(self.sheet_name(*slot))
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
                xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
                <Application>The Litchi Rust Library</Application>
                <DocSecurity>0</DocSecurity>
                <ScaleCrop>false</ScaleCrop>
                <HeadingPairs>
                    <vt:vector size="2" baseType="variant">
                        <vt:variant>
                            <vt:lpstr>Sheet</vt:lpstr>
                        </vt:variant>
                        <vt:variant>
                            <vt:i4>{}</vt:i4>
                        </vt:variant>
                    </vt:vector>
                </HeadingPairs>
                <TitlesOfParts>
                    <vt:vector size="{}" baseType="lpstr">{}</vt:vector>
                </TitlesOfParts>
                <Company></Company>
                <LinksUpToDate>false</LinksUpToDate>
                <SharedDoc>false</SharedDoc>
                <HyperlinksChanged>false</HyperlinksChanged>
                <AppVersion>14.0000</AppVersion>
            </Properties>"#,
            sheet_count,
            sheet_count,
            sheet_names
        )
    }

    /// Create core.xml content (Core Properties)
    fn create_core_xml(&self) -> String {
        // Get current timestamp in W3CDTF format
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap();
        let timestamp = format_w3cdtf(now.as_secs());

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <cp:coreProperties
                    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
                    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
                    <dc:creator>The Litchi Rust Library</dc:creator>
                    <cp:lastModifiedBy>The Litchi Rust Library</cp:lastModifiedBy>
                    <dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
                    <dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
                </cp:coreProperties>"#,
            timestamp,
            timestamp
        )
    }

    /// Add theme (REQUIRED by Excel to open file)
    fn add_theme(&self, package: &mut OpcPackage) -> XlsbResult<()> {
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
    fn create_minimal_theme(&self) -> &'static str {
        xml_minifier::minified_xml!("../resources/theme/theme1.xml")
    }

    /// Add workbook part to the package
    fn add_workbook_part(
        &self,
        package: &mut OpcPackage,
        formula_sheet_ranges: &[(u32, u32)],
    ) -> XlsbResult<()> {
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
                    XlsbSheetSlot::Worksheet(index) => {
                        rels.get_or_add(
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                            &format!("worksheets/sheet{}.bin", index + 1),
                        );
                    },
                    XlsbSheetSlot::ChartSheet(index) => {
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

    /// Write workbook structure.
    ///
    /// The record order is based on the minimal SheetJS `write_wb_bin`
    /// implementation and [MS-XLSB] examples:
    ///
    /// ```text
    /// BrtBeginBook (0x0083)
    /// BrtFileVersion (0x0080)
    /// BrtWbProp (0x0099)
    /// [BrtBeginBookViews/BrtBookView/BrtEndBookViews]
    /// BrtBeginBundleShs / BrtBundleSh / BrtEndBundleShs (0x008F / 0x009C / 0x0090)
    /// [BrtBeginPivotCacheIDs / BrtBeginPivotCacheID / BrtEndPivotCacheID / BrtEndPivotCacheIDs]
    /// BrtBeginExternals / BrtSupSelf / BrtExternSheet / BrtEndExternals
    /// [BrtCalcProp]
    /// BrtEndBook (0x0084)
    /// ```
    ///
    /// The book views and calculation properties are currently written with a
    /// single default view and sensible defaults for calculation settings.
    fn write_workbook<W: Write>(
        &self,
        writer: &mut Writer<W>,
        formula_sheet_ranges: &[(u32, u32)],
        pivot_cache_rel_ids: &[(u32, String)],
        external_link_rel_ids: &[String],
    ) -> XlsbResult<()> {
        // BrtBeginBook
        writer.write_record(kind::BEGIN_BOOK, &[])?;

        // BrtFileVersion - required by Excel
        self.write_file_version(writer)?;

        // BrtWbProp - basic workbook properties
        self.write_workbook_properties(writer)?;

        // Optional book views. We currently always emit a single default view
        // similar to SheetJS. This is small and helps some consumers which
        // expect explicit book view records.
        self.write_book_views(writer)?;

        // BrtBeginBundleShs / BrtBundleSh / BrtEndBundleShs - sheet metadata
        self.write_bundle_sheets(writer)?;

        // PivotCache identifiers, if any caches were attached.
        Self::write_pivot_cache_ids(writer, pivot_cache_rel_ids)?;

        // EXTERNALS block with self-references, mirroring SheetJS and
        // [MS-XLSB] examples. This creates a minimal but fully valid
        // extern sheet table for the workbook.
        self.write_externals(writer, formula_sheet_ranges, external_link_rel_ids)?;

        // Defined names (named ranges), if any.
        self.write_named_ranges(writer)?;

        // Basic calculation properties describing recalc behavior and
        // numerical tolerance. This is tiny and follows the spec example
        // values, so we emit it unconditionally.
        self.write_calc(writer)?;

        // BrtEndBook
        writer.write_record(kind::END_BOOK, &[])?;

        Ok(())
    }

    /// Write the PivotCache ID collection (BrtBeginPivotCacheIDs,
    /// MS-XLSB 2.4.170): one BrtBeginPivotCacheID record per attached cache,
    /// pairing the workbook cache identifier (`idSx`) with the relationship
    /// ID of its PivotCache Definition part.
    fn write_pivot_cache_ids<W: Write>(
        writer: &mut Writer<W>,
        pivot_cache_rel_ids: &[(u32, String)],
    ) -> XlsbResult<()> {
        if pivot_cache_rel_ids.is_empty() {
            return Ok(());
        }
        writer.write_record(kind::BEGIN_PIVOT_CACHE_IDS, &[])?;
        for (cache_id, rel_id) in pivot_cache_rel_ids {
            let mut data = Vec::with_capacity(rel_id.len() * 2 + 8);
            let mut temp_writer = Writer::new(&mut data);
            temp_writer.write_u32(*cache_id)?;
            temp_writer.write_wide_string(rel_id)?;
            writer.write_record(kind::BEGIN_PIVOT_CACHE_ID, &data)?;
            writer.write_record(kind::END_PIVOT_CACHE_ID, &[])?;
        }
        writer.write_record(kind::END_PIVOT_CACHE_IDS, &[])?;
        Ok(())
    }

    /// Write file version record (BrtFileVersion)
    /// This is REQUIRED for Excel to open the file
    fn write_file_version<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        // Build structure per spec example (48 bytes total):
        // guidCodeName (16 zero bytes), stAppName ("xl"), stLastEdited ("4"),
        // stLowestEdited ("4"), stRupBuild ("4505")
        let mut data = Vec::with_capacity(48);
        let mut w = Writer::new(&mut data);

        // GUID (16 bytes of zeros)
        w.write_u32(0)?;
        w.write_u32(0)?;
        w.write_u32(0)?;
        w.write_u32(0)?;

        // stAppName: "xl"
        w.write_wide_string("xl")?;
        // stLastEdited: "4"
        w.write_wide_string("4")?;
        // stLowestEdited: "4"
        w.write_wide_string("4")?;
        // stRupBuild: "4505"
        w.write_wide_string("4505")?;

        writer.write_record(kind::FILE_VERSION, &data)?;
        Ok(())
    }

    /// Write workbook properties (BrtWbProp)
    fn write_workbook_properties<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        // Flags (4 bytes). We currently only support the 1904 date system
        // bit, mirroring the minimal SheetJS implementation:
        //   bit 0 (0x0000_0001) = f1904 (date1904)
        let mut flags: u32 = 0;
        if self.is_1904 {
            flags |= 0x0000_0001;
        }
        temp_writer.write_u32(flags)?;

        // Reserved/unused DWORD (4 bytes), set to 0.
        temp_writer.write_u32(0)?;

        // Code name (XLWideString). Use the standard VBA code name
        // "ThisWorkbook" as SheetJS and Excel commonly do.
        temp_writer.write_wide_string("ThisWorkbook")?;

        writer.write_record(kind::WORKBOOK_PROP, &data)?;
        Ok(())
    }

    /// Write book views (REQUIRED by Excel)
    fn write_book_views<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        writer.write_record(kind::BEGIN_BOOK_VIEWS, &[])?;

        // Write one default book view
        let mut view_data = Vec::new();
        let mut temp_writer = Writer::new(&mut view_data);

        // xWn (4), yWn (4), dxWn (4), dyWn (4)
        temp_writer.write_u32(0)?; // xWn
        temp_writer.write_u32(0)?; // yWn
        temp_writer.write_u32(0x00004E20)?; // dxWn (width)
        temp_writer.write_u32(0x00002710)?; // dyWn (height)

        // iTabRatio (4): 0 means auto
        temp_writer.write_u32(0)?;
        // itabFirst (4): first visible bundle sheet index
        temp_writer.write_u32(0)?;
        // itabCur (4): active sheet index
        temp_writer.write_u32(0)?;

        // Flags (1 byte) - D/E/F bits set for scrollbars and tabs
        temp_writer.write_u8(0x78)?; // Total: 7*4 + 1 = 29 bytes

        writer.write_record(kind::BOOK_VIEW, &view_data)?;

        writer.write_record(kind::END_BOOK_VIEWS, &[])?;
        Ok(())
    }

    /// Write bundle sheets in workbook order.
    fn write_bundle_sheets<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        writer.write_record(kind::BEGIN_BUNDLE_SHS, &[])?;

        for (i, slot) in self.sheet_order.iter().copied().enumerate() {
            let mut sheet_data = Vec::new();
            let mut temp_writer = Writer::new(&mut sheet_data);

            let state = match slot {
                XlsbSheetSlot::Worksheet(_) => 0,
                XlsbSheetSlot::ChartSheet(index) => match self.chart_sheets[index].metadata().state
                {
                    crate::xlsb::chartsheet::XlsbChartSheetState::Visible => 0,
                    crate::xlsb::chartsheet::XlsbChartSheetState::Hidden => 1,
                    crate::xlsb::chartsheet::XlsbChartSheetState::VeryHidden => 2,
                },
            };
            temp_writer.write_u32(state)?;
            // itabID (u32): unique sheet id (1-based)
            temp_writer.write_u32(u32::try_from(i + 1).map_err(|_| {
                XlsbError::InvalidFormula("sheet identifier overflow".to_string())
            })?)?;
            // RelID (XLWideString): rIdN
            temp_writer.write_wide_string(&format!("rId{}", i + 1))?;
            // strName (XLWideString): sheet name
            temp_writer.write_wide_string(self.sheet_name(slot))?;

            writer.write_record(kind::BUNDLE_SH, &sheet_data)?;
        }

        writer.write_record(kind::END_BUNDLE_SHS, &[])?;
        Ok(())
    }

    /// Write calculation properties (CALC_PROP, 0x009D)
    ///
    /// Spec example fields and order
    fn write_calc<W: Write>(&self, writer: &mut Writer<W>) -> XlsbResult<()> {
        writer.write_header(kind::CALC_PROP, calc::LEN)?;
        calc::write(&self.calc, writer)?;
        Ok(())
    }

    /// Write externals section (self-references)
    ///
    /// Based on SheetJS implementation: always writes BrtSupSelf with BrtExternSheet
    /// This creates self-references for the workbook and all sheets.
    fn write_externals<W: Write>(
        &self,
        writer: &mut Writer<W>,
        formula_sheet_ranges: &[(u32, u32)],
        external_link_rel_ids: &[String],
    ) -> XlsbResult<()> {
        // BrtBeginExternals - no data
        writer.write_record(kind::BEGIN_EXTERNALS, &[])?;

        // BrtSupSelf - no data
        writer.write_record(kind::SUP_SELF, &[])?;

        for relationship_id in external_link_rel_ids {
            let mut data = Vec::with_capacity(4 + relationship_id.len() * 2);
            Writer::new(&mut data).write_wide_string(relationship_id)?;
            writer.write_record(kind::SUP_BOOK_SRC, &data)?;
        }

        // BrtExternSheet - self-references data
        let mut data = Vec::new();
        let mut temp_writer = Writer::new(&mut data);

        let sheet_count = self.sheet_order.len();

        // Total count: workbook and #REF entries, single-sheet entries, then
        // the distinct multi-sheet ranges referenced by formulas.
        let entry_count = sheet_count
            .checked_add(2)
            .and_then(|count| count.checked_add(formula_sheet_ranges.len()))
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtExternSheet entry count overflow".to_string(),
                )
            })?;
        if entry_count >= 65_536 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtExternSheet entry count {entry_count} exceeds 65,535"
            )));
        }
        temp_writer.write_u32(u32::try_from(entry_count).map_err(|_| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "BrtExternSheet entry count overflow".to_string(),
            )
        })?)?;

        // First entry: workbook-level reference (0, -2, -2)
        temp_writer.write_u32(0)?;
        temp_writer.write_i32(-2)?;
        temp_writer.write_i32(-2)?;

        // Second entry: #REF! (0, -1, -1)
        temp_writer.write_u32(0)?;
        temp_writer.write_i32(-1)?;
        temp_writer.write_i32(-1)?;

        // Then for each sheet: (0, sheet_index, sheet_index)
        for i in 0..sheet_count {
            temp_writer.write_u32(0)?;
            temp_writer.write_i32(i as i32)?;
            temp_writer.write_i32(i as i32)?;
        }

        for &(first_sheet, last_sheet) in formula_sheet_ranges {
            if last_sheet < first_sheet
                || usize::try_from(last_sheet)
                    .ok()
                    .is_none_or(|last_sheet| last_sheet >= sheet_count)
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "invalid formula sheet range {first_sheet}..={last_sheet}"
                )));
            }
            temp_writer.write_u32(0)?;
            temp_writer.write_i32(i32::try_from(first_sheet).map_err(|_| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "first formula sheet index overflow".to_string(),
                )
            })?)?;
            temp_writer.write_i32(i32::try_from(last_sheet).map_err(|_| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "last formula sheet index overflow".to_string(),
                )
            })?)?;
        }

        writer.write_record(kind::EXTERN_SHEET, &data)?;

        // BrtEndExternals - no data
        writer.write_record(kind::END_EXTERNALS, &[])?;

        Ok(())
    }

    /// Add worksheet parts to the package
    fn add_worksheet_parts(&mut self, package: &mut OpcPackage) -> XlsbResult<Vec<(u32, u32)>> {
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
                XlsbSheetSlot::Worksheet(worksheet_index) => Some((*worksheet_index, sheet_index)),
                XlsbSheetSlot::ChartSheet(_) => None,
            })
            .collect::<std::collections::HashMap<_, _>>();
        let defined_names = self
            .named_ranges
            .iter()
            .map(|named_range| FormulaDefinedName {
                name: named_range.name.clone(),
                sheet_id: named_range.sheet_id,
            })
            .collect::<Vec<_>>();
        let formula_sheet_ranges = std::cell::RefCell::new(Vec::new());
        let mut next_table_index = 1usize;
        let mut next_drawing_index = 1usize;
        let mut next_chart_index = 1usize;
        let mut next_image_index = 1usize;
        let mut next_pivot_table_index = 1usize;
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
            let binary_index_uri = PackURI::new(format!("/xl/worksheets/{}", binary_index_name))?;
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
                crate::xlsb::comments::write_comments(
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
                for table in worksheet.tables() {
                    let table_name = format!("tables/table{next_table_index}.bin");
                    next_table_index += 1;
                    rel_ids.push(sheet_part.relate_to(&format!("../{table_name}"), rel::TABLE));
                    table_parts.push(BlobPart::new(
                        PackURI::new(format!("/xl/{table_name}"))?,
                        "application/vnd.ms-excel.table".to_string(),
                        crate::xlsb::table::write::write_table_part(table)?,
                    ));
                }
                worksheet.table_rel_ids = rel_ids;
            }

            // PivotTable definitions are related implicitly from their host
            // worksheet and back to the exact workbook PivotCache definition.
            let mut pivot_table_parts = Vec::new();
            for view in worksheet.pivot_table_views() {
                let cache_index = self
                    .pivot_caches
                    .iter()
                    .position(|cache| cache.id == view.cache_id())
                    .ok_or_else(|| {
                        XlsbError::InvalidFormula(format!(
                            "PivotTable view {:?} references unknown cache {}",
                            view.name(),
                            view.cache_id()
                        ))
                    })?;
                let cache_version_created = self.pivot_caches[cache_index].version_created;
                if (view.version_created() >= 3) != (cache_version_created >= 3) {
                    return Err(XlsbError::InvalidFormula(format!(
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
                        XlsbError::InvalidFormula("PivotTable part index overflow".to_string())
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
                    .collect::<XlsbResult<Vec<_>>>()?;
                let drawing_xml = crate::xlsb::drawing_write::serialize_drawing(
                    worksheet.images(),
                    &normalized_charts,
                    worksheet.shapes(),
                    worksheet.groups(),
                    worksheet.connections(),
                )?;
                let drawing_name = format!("drawing{next_drawing_index}.xml");
                next_drawing_index = next_drawing_index.checked_add(1).ok_or_else(|| {
                    XlsbError::InvalidFormula("drawing part index overflow".to_string())
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
                        XlsbError::InvalidFormula("image part index overflow".to_string())
                    })?;
                    let relationship_id =
                        part.relate_to(&format!("../media/{image_name}"), rel::IMAGE);
                    let expected_relationship_id = format!("rId{}", image_ordinal + 1);
                    if relationship_id != expected_relationship_id {
                        return Err(XlsbError::InvalidFormula(format!(
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
                    let graph =
                        crate::xlsb::chart_resources::author_chart_graph(chart, next_chart_index)?;
                    next_chart_index = next_chart_index.checked_add(1).ok_or_else(|| {
                        XlsbError::InvalidFormula("chart part index overflow".to_string())
                    })?;
                    let relationship_id =
                        part.relate_to(&format!("../charts/{chart_name}"), rel::CHART);
                    let expected_relationship_id =
                        format!("rId{}", worksheet.images().len() + chart_ordinal + 1);
                    if relationship_id != expected_relationship_id {
                        return Err(XlsbError::InvalidFormula(format!(
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
            let current_sheet = u32::try_from(worksheet_sheet_indexes[&i]).map_err(|_| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "worksheet index overflow".to_string(),
                )
            })?;
            let formula_context = FormulaCompilationContext {
                worksheet_names: &worksheet_names,
                defined_names: &defined_names,
                tables: &[],
                supporting_links: &[],
                external_sheets: &[],
                external_books: &[],
                sheet_ranges: &formula_sheet_ranges,
                current_sheet,
            };
            let compiled_formulas = worksheet.compile_contextual_formulas(&formula_context)?;
            let write_result = {
                let mut writer = Writer::new(&mut sheet_data);
                worksheet.write(&mut writer, &mut self.shared_strings)
            };
            worksheet.clear_compiled_formulas(compiled_formulas);
            write_result?;
            sheet_part.set_blob(sheet_data);

            package.add_part(Box::new(sheet_part));
            package.add_part(Box::new(binary_index_part));
            if let Some(part) = comments_part {
                package.add_part(Box::new(part));
            }
            for part in table_parts {
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

        Ok(formula_sheet_ranges.into_inner())
    }

    /// Add binary Chart Sheet streams and their standard DrawingML graphs.
    fn add_chart_sheet_parts(&self, package: &mut OpcPackage) -> XlsbResult<()> {
        let mut next_drawing_index = self
            .worksheets
            .iter()
            .filter(|sheet| sheet.has_drawing_objects())
            .count()
            .checked_add(1)
            .ok_or_else(|| XlsbError::InvalidFormula("drawing part index overflow".to_string()))?;
        let mut next_chart_index = self
            .worksheets
            .iter()
            .try_fold(1usize, |next, sheet| next.checked_add(sheet.charts().len()))
            .ok_or_else(|| XlsbError::InvalidFormula("chart part index overflow".to_string()))?;

        for (index, sheet) in self.chart_sheets.iter().enumerate() {
            sheet.validate()?;
            let drawing_name = format!("drawing{next_drawing_index}.xml");
            next_drawing_index = next_drawing_index.checked_add(1).ok_or_else(|| {
                XlsbError::InvalidFormula("drawing part index overflow".to_string())
            })?;
            let chart_index = next_chart_index;
            let chart_name = format!("chart{chart_index}.xml");
            next_chart_index = next_chart_index.checked_add(1).ok_or_else(|| {
                XlsbError::InvalidFormula("chart part index overflow".to_string())
            })?;
            let normalized_chart = Self::normalized_pivot_chart(
                sheet.chart(),
                sheet.name(),
                &self.authored_pivot_tables(),
            )?;
            let graph =
                crate::xlsb::chart_resources::author_chart_graph(&normalized_chart, chart_index)?;

            let mut chart_sheet_part = BlobPart::new(
                PackURI::new(format!("/xl/chartsheets/sheet{}.bin", index + 1))?,
                "application/vnd.ms-excel.chartsheet".to_string(),
                Vec::new(),
            );
            let drawing_rel_id =
                chart_sheet_part.relate_to(&format!("../drawings/{drawing_name}"), rel::DRAWING);
            if drawing_rel_id != "rId1" {
                return Err(XlsbError::InvalidFormula(format!(
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
            chart_sheet_part.set_blob(super::chartsheet::write_chart_sheet(
                sheet,
                &drawing_rel_id,
                printer_rel_id.as_deref(),
            )?);

            let mut drawing_part = BlobPart::new(
                PackURI::new(format!("/xl/drawings/{drawing_name}"))?,
                ct::OFC_DRAWING.to_string(),
                crate::xlsb::drawing_write::serialize_chart_sheet_drawing(sheet.name())?,
            );
            let chart_rel_id =
                drawing_part.relate_to(&format!("../charts/{chart_name}"), rel::CHART);
            if chart_rel_id != "rId1" {
                return Err(XlsbError::InvalidFormula(format!(
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
    fn add_shared_strings_part(&self, package: &mut OpcPackage) -> XlsbResult<()> {
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
    fn add_styles_part(&self, package: &mut OpcPackage) -> XlsbResult<()> {
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

/// Format Unix timestamp as W3CDTF (ISO 8601)
fn format_w3cdtf(secs: u64) -> String {
    // Simple conversion: seconds since 1970-01-01 to ISO 8601
    // This is a simplified version; for production, use chrono or time crate
    let days = secs / 86400;
    let year = 1970 + (days / 365);
    let day_of_year = days % 365;
    let month = ((day_of_year / 30) + 1).min(12);
    let day = ((day_of_year % 30) + 1).min(31);

    let time_of_day = secs % 86400;
    let hours = time_of_day / 3600;
    let minutes = (time_of_day % 3600) / 60;
    let seconds = time_of_day % 60;

    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}Z",
        year, month, day, hours, minutes, seconds
    )
}

impl Default for XlsbWorkbookWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::comments::Comment;
    use crate::xlsb::conditional_formatting::{
        CfRuleType, Cfvo, ColorScale, ConditionalFormatColor, ConditionalFormatting,
        ConditionalFormattingRecordKind, ConditionalFormattingRule,
        ConditionalFormattingRule14Metadata, DataBar, DataBar14, IconSet,
    };
    use crate::xlsb::data_validation::{
        DataValidation, DataValidationRecordKind, DataValidationSettings,
    };
    use crate::xlsb::{SharedStringRun, SheetProtection};
    use litchi_core::sheet::{CellValue, WorkbookTrait};
    use litchi_xlsb::calc::{Delta, Mode, Opts, Threads};
    use std::io::Cursor;

    #[test]
    fn test_create_empty_workbook() {
        let workbook = XlsbWorkbookWriter::new();
        assert_eq!(workbook.worksheet_count(), 0);
        assert!(!workbook.is_1904);
    }

    #[test]
    fn test_add_worksheet() {
        let mut workbook = XlsbWorkbookWriter::new();
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        workbook.add_worksheet(sheet);
        assert_eq!(workbook.worksheet_count(), 1);
    }

    #[test]
    fn external_links_round_trip_with_inert_package_topology() {
        use crate::xlsb::external_link::{
            XlsbDdeItem, XlsbExternalCachedValue, XlsbExternalCellLocation,
            XlsbExternalCellReference, XlsbExternalDefinedName, XlsbExternalErrorValue,
            XlsbExternalLink, XlsbExternalLinkKind, XlsbExternalNameFormula,
            XlsbExternalNameFormulaKind, XlsbExternalSheetRange, XlsbExternalValueMatrix,
            XlsbOleItem,
        };

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Host"));
        let external_formula =
            XlsbExternalNameFormula::cell_reference(XlsbExternalCellReference::new(
                XlsbExternalSheetRange::sheets(0, 0).unwrap(),
                XlsbExternalCellLocation::new(3, 2),
            ));
        workbook
            .add_external_link(
                XlsbExternalLink::workbook_with_defined_names(
                    "file:///data/Budget.xlsx",
                    vec!["Data".to_string(), "Rates".to_string()],
                    vec![
                        XlsbExternalDefinedName::new("ExchangeRate")
                            .unwrap()
                            .with_formula(external_formula)
                            .with_built_in(true)
                            .with_sheet_scope(1),
                    ],
                )
                .unwrap(),
            )
            .unwrap();
        let dde_cache = XlsbExternalValueMatrix::new(
            1,
            5,
            vec![
                XlsbExternalCachedValue::Empty,
                XlsbExternalCachedValue::Number(42.5),
                XlsbExternalCachedValue::Boolean(true),
                XlsbExternalCachedValue::Error(XlsbExternalErrorValue::NotAvailable),
                XlsbExternalCachedValue::String("Ready".to_string()),
            ],
        )
        .unwrap();
        workbook
            .add_external_link(
                XlsbExternalLink::dde_with_items(
                    "Excel",
                    "System",
                    vec![
                        XlsbDdeItem::new("StatusItem")
                            .unwrap()
                            .with_advise(true)
                            .with_picture(true)
                            .with_cached_values(dde_cache),
                    ],
                )
                .unwrap(),
            )
            .unwrap();
        workbook
            .add_external_link(
                XlsbExternalLink::ole_with_items(
                    "file:///data/Model.xlsx",
                    "Excel.Sheet.12",
                    vec![
                        XlsbOleItem::new("ModelItem")
                            .unwrap()
                            .with_advise(true)
                            .with_icon(true)
                            .with_cached_values(
                                XlsbExternalValueMatrix::new(
                                    1,
                                    1,
                                    vec![XlsbExternalCachedValue::Number(7.0)],
                                )
                                .unwrap(),
                            ),
                    ],
                )
                .unwrap(),
            )
            .unwrap();

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let bytes = output.into_inner();
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let workbook_part = package
            .get_part(&PackURI::new("/xl/workbook.bin").unwrap())
            .unwrap();
        let external_relationships = workbook_part
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == rel::EXTERNAL_LINK)
            .collect::<Vec<_>>();
        assert_eq!(external_relationships.len(), 3);
        let relationships = (1..=3)
            .map(|index| {
                let expected =
                    PackURI::new(format!("/xl/externalLinks/externalLink{index}.bin")).unwrap();
                external_relationships
                    .iter()
                    .copied()
                    .find(|relationship| relationship.target_partname().unwrap() == expected)
                    .expect("external-link relationship missing")
            })
            .collect::<Vec<_>>();
        let mut support_relationship_ids = Vec::new();
        for record in litchi_xlsb::raw::Records::new(workbook_part.blob()) {
            let record = record.unwrap();
            if record.kind() == kind::SUP_BOOK_SRC {
                let (relationship_id, consumed) =
                    crate::xlsb::records::decode_string(record.payload()).unwrap();
                assert_eq!(consumed, record.payload().len());
                support_relationship_ids.push(relationship_id);
            }
        }
        assert_eq!(
            support_relationship_ids,
            relationships
                .iter()
                .map(|relationship| relationship.r_id().to_string())
                .collect::<Vec<_>>()
        );
        for (index, relationship) in relationships.iter().enumerate() {
            assert!(!relationship.is_external());
            assert_eq!(
                relationship.target_partname().unwrap(),
                PackURI::new(format!("/xl/externalLinks/externalLink{}.bin", index + 1)).unwrap()
            );
        }
        let workbook_link = package
            .get_part(&relationships[0].target_partname().unwrap())
            .unwrap();
        assert_eq!(workbook_link.rels().len(), 1);
        assert!(
            workbook_link
                .rels()
                .iter()
                .all(|relationship| relationship.is_external()
                    && relationship.reltype() == rel::EXTERNAL_LINK_PATH)
        );
        let dde_link = package
            .get_part(&relationships[1].target_partname().unwrap())
            .unwrap();
        assert!(dde_link.rels().is_empty());
        let ole_link = package
            .get_part(&relationships[2].target_partname().unwrap())
            .unwrap();
        assert_eq!(ole_link.rels().len(), 1);
        assert!(
            ole_link
                .rels()
                .iter()
                .all(|relationship| relationship.is_external()
                    && relationship.reltype() == rel::OLE_OBJECT)
        );

        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(bytes)).unwrap();
        assert_eq!(reader.external_link_count(), 3);
        assert_eq!(reader.external_link_iter().len(), 3);
        assert_eq!(reader.external_link(1).unwrap().dde_topic(), Some("System"));
        let links = reader.external_links();
        assert_eq!(links.len(), 3);
        assert_eq!(links[0].kind(), XlsbExternalLinkKind::Workbook);
        assert_eq!(links[0].source(), "file:///data/Budget.xlsx");
        assert_eq!(links[0].sheet_names(), ["Data", "Rates"]);
        let defined_name = &links[0].defined_names()[0];
        assert_eq!(defined_name.name(), "ExchangeRate");
        assert!(defined_name.is_built_in());
        assert_eq!(defined_name.scope_sheet_index(), Some(1));
        assert_eq!(
            defined_name.formula().unwrap().kind(),
            XlsbExternalNameFormulaKind::CellReference
        );
        assert_eq!(
            defined_name.formula().unwrap().tokens(),
            [0x3A, 0, 0, 0, 0, 3, 0, 2, 0]
        );
        assert_eq!(links[1].kind(), XlsbExternalLinkKind::Dde);
        assert_eq!(links[1].source(), "Excel");
        assert_eq!(links[1].dde_topic(), Some("System"));
        let dde_item = &links[1].dde_items()[0];
        assert_eq!(dde_item.name(), "StatusItem");
        assert!(dde_item.wants_advise());
        assert!(dde_item.wants_picture());
        assert_eq!(dde_item.cached_values().unwrap().rows(), 1);
        assert_eq!(dde_item.cached_values().unwrap().columns(), 5);
        assert_eq!(
            dde_item.cached_values().unwrap().values(),
            [
                XlsbExternalCachedValue::Empty,
                XlsbExternalCachedValue::Number(42.5),
                XlsbExternalCachedValue::Boolean(true),
                XlsbExternalCachedValue::Error(XlsbExternalErrorValue::NotAvailable),
                XlsbExternalCachedValue::String("Ready".to_string()),
            ]
        );
        assert_eq!(links[2].kind(), XlsbExternalLinkKind::Ole);
        assert_eq!(links[2].source(), "file:///data/Model.xlsx");
        assert_eq!(links[2].ole_program_id(), Some("Excel.Sheet.12"));
        let ole_item = &links[2].ole_items()[0];
        assert_eq!(ole_item.name(), "ModelItem");
        assert!(ole_item.wants_advise());
        assert!(ole_item.displays_as_icon());
        assert_eq!(
            ole_item.cached_values().unwrap().values(),
            [XlsbExternalCachedValue::Number(7.0)]
        );
    }

    #[test]
    fn external_link_constructors_refuse_malformed_metadata() {
        use crate::xlsb::external_link::{
            XlsbDdeItem, XlsbExternalCachedValue, XlsbExternalDefinedName, XlsbExternalLink,
            XlsbExternalNameFormula, XlsbExternalValueMatrix,
        };

        assert!(XlsbExternalLink::workbook("", Vec::new(), Vec::new()).is_err());
        assert!(
            XlsbExternalLink::workbook(
                "Book.xlsx",
                vec!["Data".to_string(), "data".to_string()],
                Vec::new(),
            )
            .is_err()
        );
        assert!(XlsbExternalLink::dde("Excel", "", vec!["Item".to_string()]).is_err());
        assert!(
            XlsbExternalLink::ole("Model.xlsx", "Excel.Sheet", vec!["A1".to_string()],).is_err()
        );
        assert!(XlsbExternalNameFormula::from_tokens(vec![0x3A, 0]).is_err());
        assert!(
            XlsbExternalValueMatrix::new(2, 2, vec![XlsbExternalCachedValue::Number(1.0)]).is_err()
        );
        assert!(
            XlsbExternalValueMatrix::new(1, 1, vec![XlsbExternalCachedValue::Number(-0.0)])
                .is_err()
        );
        assert!(
            XlsbExternalValueMatrix::new(
                1,
                1,
                vec![XlsbExternalCachedValue::Number(f64::from_bits(1))]
            )
            .is_err()
        );
        assert!(
            XlsbExternalValueMatrix::new(
                1,
                1,
                vec![XlsbExternalCachedValue::String(String::new())]
            )
            .is_ok()
        );
        let invalid_formula =
            XlsbExternalNameFormula::from_tokens(vec![0x3A, 2, 0, 2, 0, 0, 0, 0, 0]).unwrap();
        assert!(
            XlsbExternalLink::workbook_with_defined_names(
                "Book.xlsx",
                vec!["OnlySheet".to_string()],
                vec![
                    XlsbExternalDefinedName::new("BadScope")
                        .unwrap()
                        .with_formula(invalid_formula)
                ],
            )
            .is_err()
        );
        assert!(
            XlsbExternalLink::dde_with_items(
                "Excel",
                "System",
                vec![
                    XlsbDdeItem::new("NotStdDocumentName")
                        .unwrap()
                        .with_ole_support(true)
                        .with_cached_values(
                            XlsbExternalValueMatrix::new(
                                1,
                                1,
                                vec![XlsbExternalCachedValue::Empty]
                            )
                            .unwrap()
                        )
                ],
            )
            .is_err()
        );
    }

    #[test]
    fn chart_sheet_metadata_chart_and_printer_settings_round_trip_in_sheet_order() {
        use crate::xlsb::chartsheet::{
            XlsbChartSheetColor, XlsbChartSheetColorType, XlsbChartSheetPageSetup,
            XlsbChartSheetProtection, XlsbChartSheetState, XlsbChartSheetView,
        };
        use crate::xlsb::worksheet::XlsbStrongProtection;
        use crate::xlsx::{ChartAnchor, WorksheetChart};

        let chart = WorksheetChart::bar_chart_with_cache(
            "Sales",
            "Data!$A$2:$A$3",
            &["North", "South"],
            "Data!$B$2:$B$3",
            &[42.0, 55.0],
            ChartAnchor::new(0, 0, 10, 20),
        )
        .unwrap();
        let mut chart_sheet = MutableXlsbChartSheet::new("Sales Chart", chart);
        {
            let metadata = chart_sheet.metadata_mut();
            metadata.state = XlsbChartSheetState::Hidden;
            metadata.code_name = "ChartCode".to_string();
            metadata.published = true;
            metadata.tab_color = XlsbChartSheetColor {
                valid_rgb: true,
                color_type: XlsbChartSheetColorType::Rgb,
                index: 0,
                tint: -100,
                rgba: [0x44, 0x72, 0xc4, 0xff],
            };
            metadata.views = vec![XlsbChartSheetView {
                selected: true,
                scale: 125,
                workbook_view_index: 0,
            }];
            metadata.protection = Some(XlsbChartSheetProtection {
                password_verifier: 0x1234,
                locked: true,
                objects: false,
            });
            metadata.strong_protection = Some(XlsbStrongProtection {
                spin_count: 100_000,
                hash: vec![7; 64],
                salt: vec![3; 16],
                algorithm: "SHA-512".to_string(),
            });
        }
        chart_sheet
            .set_page_setup(
                XlsbChartSheetPageSetup {
                    paper_size: 9,
                    horizontal_resolution: 600,
                    vertical_resolution: 600,
                    copies: 2,
                    page_start: 4,
                    landscape: true,
                    black_and_white: true,
                    use_default_orientation: false,
                    use_page_start: true,
                    draft: false,
                    printer_settings_rel_id: "caller-id-is-replaced".to_string(),
                },
                vec![1, 2, 3, 4],
            )
            .unwrap();

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data"));
        workbook.add_chart_sheet(chart_sheet).unwrap();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Tail"));
        assert_eq!(workbook.chart_sheet_count(), 1);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let bytes = output.into_inner();
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(bytes)).unwrap();
        assert_eq!(
            reader.worksheet_names(),
            &[
                "Data".to_string(),
                "Sales Chart".to_string(),
                "Tail".to_string()
            ]
        );
        let parsed = reader.chart_sheet(1).expect("chart sheet missing");
        assert_eq!(parsed.state, XlsbChartSheetState::Hidden);
        assert_eq!(parsed.code_name, "ChartCode");
        assert!(parsed.published);
        assert_eq!(parsed.tab_color.rgba, [0x44, 0x72, 0xc4, 0xff]);
        assert_eq!(parsed.views[0].scale, 125);
        assert_eq!(parsed.protection.unwrap().password_verifier, 0);
        assert_eq!(
            parsed.strong_protection.as_ref().unwrap().algorithm,
            "SHA-512"
        );
        let page_setup = parsed.page_setup.as_ref().unwrap();
        assert_eq!(page_setup.copies, 2);
        assert_eq!(page_setup.printer_settings_rel_id, "rId2");

        let drawing = reader.sheet_drawing(1).expect("chart drawing missing");
        assert_eq!(drawing.drawing.anchors.len(), 1);
        assert_eq!(drawing.charts.len(), 1);
        let printer = package
            .get_part(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
            .unwrap();
        assert_eq!(printer.blob(), &[1, 2, 3, 4]);
    }

    #[test]
    fn chart_sheet_validation_is_lossless_or_refuse() {
        use crate::xlsb::chartsheet::{XlsbChartSheetColor, XlsbChartSheetColorType};
        use crate::xlsx::{ChartAnchor, WorksheetChart};

        let chart = WorksheetChart::bar_chart(
            "T",
            "Data!$A$1:$A$2",
            "Data!$B$1:$B$2",
            ChartAnchor::new(0, 0, 5, 5),
        )
        .unwrap();
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data"));
        workbook
            .add_chart_sheet(MutableXlsbChartSheet::new("Chart", chart.clone()))
            .unwrap();
        assert!(
            workbook
                .add_chart_sheet(MutableXlsbChartSheet::new("chart", chart.clone()))
                .is_err()
        );

        let mut invalid = MutableXlsbChartSheet::new("Invalid", chart.clone());
        invalid.metadata_mut().views[0].scale = 401;
        assert!(workbook.add_chart_sheet(invalid).is_err());

        let mut invalid_color = MutableXlsbChartSheet::new("Invalid Color", chart);
        invalid_color.metadata_mut().tab_color = XlsbChartSheetColor {
            valid_rgb: false,
            color_type: XlsbChartSheetColorType::Indexed,
            index: 0x52,
            tint: 0,
            rgba: [0; 4],
        };
        assert!(workbook.add_chart_sheet(invalid_color).is_err());
    }

    #[test]
    fn test_set_date_system() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.set_date_system(true);
        assert!(workbook.is_1904);
    }

    #[test]
    fn calc_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let properties = workbook.calc_mut();
        properties
            .set_mode(Mode::Manual)
            .set_iters(25)
            .set_delta(Delta::new(0.000_01).unwrap())
            .set_threads(Threads::new(4).unwrap());
        properties
            .set_opt(
                Opts::ITERATE | Opts::USER_THREADS | Opts::FULL_ON_LOAD,
                true,
            )
            .unwrap();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));

        let expected = workbook.calc().clone();
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        assert_eq!(reader.calc(), &expected);
    }

    #[test]
    fn connections_round_trip_through_save_and_read() {
        use crate::xlsb::connections::*;

        let connections = XlsbConnections {
            connections: vec![
                XlsbConnection {
                    connection_id: 42,
                    source_type: XlsbConnectionSourceType::Odbc,
                    name: "Warehouse".to_string(),
                    refresh_interval_minutes: 30,
                    background_query: true,
                    credential_method: Some(XlsbCredentialMethod::Integrated),
                    properties: XlsbConnectionProperties::Database(XlsbDbProperties {
                        command_type: XlsbCommandType::Sql,
                        connection_string: "Driver={SQL Server};Server=db".to_string(),
                        command: Some("SELECT * FROM T".to_string()),
                        server_command: None,
                    }),
                    ..XlsbConnection::default()
                },
                XlsbConnection {
                    connection_id: 9,
                    source_type: XlsbConnectionSourceType::Web,
                    name: "Web Query".to_string(),
                    properties: XlsbConnectionProperties::Web(XlsbWebProperties {
                        html_format: XlsbHtmlFormat::All,
                        url: Some("https://example.test/q".to_string()),
                        ..XlsbWebProperties::default()
                    }),
                    web_tables: vec![XlsbWebTableItem::Index(1)],
                    ..XlsbConnection::default()
                },
            ],
        };

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        workbook.set_connections(connections.clone()).unwrap();
        // Validation: zero id, duplicate id, duplicate name (case-insensitive).
        assert!(
            workbook
                .set_connections(XlsbConnections {
                    connections: vec![XlsbConnection {
                        connection_id: 0,
                        name: "bad".to_string(),
                        ..XlsbConnection::default()
                    }],
                })
                .is_err()
        );
        assert!(
            workbook
                .set_connections(XlsbConnections {
                    connections: vec![
                        XlsbConnection {
                            connection_id: 5,
                            name: "a".to_string(),
                            ..XlsbConnection::default()
                        },
                        XlsbConnection {
                            connection_id: 5,
                            name: "b".to_string(),
                            ..XlsbConnection::default()
                        },
                    ],
                })
                .is_err()
        );
        assert!(
            workbook
                .set_connections(XlsbConnections {
                    connections: vec![
                        XlsbConnection {
                            connection_id: 5,
                            name: "Dup".to_string(),
                            ..XlsbConnection::default()
                        },
                        XlsbConnection {
                            connection_id: 6,
                            name: "dup".to_string(),
                            ..XlsbConnection::default()
                        },
                    ],
                })
                .is_err()
        );

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let parsed = reader.connections().expect("connections part missing");
        assert_eq!(parsed, &connections);
        assert_eq!(parsed.by_id(42).unwrap().name, "Warehouse");
        assert!(parsed.by_name("Web Query").is_some());
    }

    #[test]
    fn structured_tables_round_trip_through_save_and_read() {
        use crate::xlsb::table::{
            XlsbTable, XlsbTableColumn, XlsbTableFormula, XlsbTableRange, XlsbTableStyleInfo,
            XlsbTableTotalsRowFunction, XlsbTableType,
        };

        let table = XlsbTable {
            id: 3,
            name: Some("SalesTable".to_string()),
            display_name: Some("SalesTable".to_string()),
            range: XlsbTableRange {
                first_row: 0,
                last_row: 2,
                first_column: 0,
                last_column: 1,
            },
            table_type: XlsbTableType::Range,
            header_row_count: 1,
            columns: vec![
                XlsbTableColumn {
                    id: 1,
                    name: Some("Region".to_string()),
                    ..XlsbTableColumn::default()
                },
                XlsbTableColumn {
                    id: 2,
                    name: Some("Amount".to_string()),
                    totals_row_function: XlsbTableTotalsRowFunction::Sum,
                    calculated_column_formula: Some(XlsbTableFormula {
                        array: false,
                        tokens: vec![0x1E, 0x02],
                        extra: Vec::new(),
                    }),
                    ..XlsbTableColumn::default()
                },
            ],
            style_info: Some(XlsbTableStyleInfo {
                name: Some("TableStyleMedium2".to_string()),
                show_first_column: false,
                show_last_column: false,
                show_row_stripes: true,
                show_column_stripes: false,
            }),
            ..XlsbTable::default()
        };

        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Sheet1");
        sheet.set_cell(0, 0, "Region");
        sheet.set_cell(0, 1, "Amount");
        sheet.set_cell(1, 0, "North");
        sheet.set_cell(1, 1, 42.5);
        sheet.add_table(table.clone()).unwrap();
        // Validation: missing display name, inverted range, width mismatch,
        // duplicate id.
        assert!(
            sheet
                .add_table(XlsbTable {
                    id: 9,
                    range: table.range,
                    ..XlsbTable::default()
                })
                .is_err()
        );
        assert!(
            sheet
                .add_table(XlsbTable {
                    id: 9,
                    display_name: Some("Bad".to_string()),
                    range: XlsbTableRange {
                        first_row: 5,
                        last_row: 2,
                        first_column: 0,
                        last_column: 0,
                    },
                    ..XlsbTable::default()
                })
                .is_err()
        );
        assert!(
            sheet
                .add_table(XlsbTable {
                    id: 9,
                    display_name: Some("Bad".to_string()),
                    range: table.range,
                    columns: vec![XlsbTableColumn::default()],
                    ..XlsbTable::default()
                })
                .is_err()
        );
        let mut duplicate = table.clone();
        duplicate.display_name = Some("Other".to_string());
        assert!(sheet.add_table(duplicate).is_err());
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let tables = reader.structured_tables();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].0, 0);
        assert_eq!(tables[0].1, table);
        assert_eq!(reader.tables_on_sheet(0).len(), 1);
        assert!(reader.tables_on_sheet(1).is_empty());
    }

    #[test]
    fn comments_survive_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Notes");
        let mut first = Comment::new(2, 3, "Author".to_string(), "formatted note".to_string());
        first.runs = vec![SharedStringRun {
            character_index: 0,
            font_id: 0,
        }];
        first.guid = [7; 16];
        sheet.add_comment(first);
        sheet.add_comment(Comment::new(
            4,
            1,
            "Author".to_string(),
            "second note".to_string(),
        ));
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();
        assert_eq!(worksheet.comments().len(), 2);
        assert_eq!(worksheet.comments()[0].text, "formatted note");
        assert_eq!(worksheet.comments()[0].runs.len(), 1);
        assert_eq!(worksheet.comments()[0].guid, [7; 16]);
        assert_eq!(worksheet.comments()[1].author, "Author");
    }

    #[test]
    fn row_and_column_formatting_survive_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Layout");
        sheet.set_cell(3, 2, "value");
        sheet.set_column_width(2, 18.25);
        sheet.set_column_hidden(2, true);
        sheet.set_column_best_fit(2, true);
        sheet.set_row_height(3, 24.5);
        sheet.set_row_hidden(3, true);
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();

        assert_eq!(worksheet.column_infos().len(), 1);
        let column = &worksheet.column_infos()[0];
        assert_eq!((column.first_column, column.last_column), (2, 2));
        assert_eq!(column.width, 18.25);
        assert!(column.user_set_width);
        assert!(column.hidden);
        assert!(column.best_fit);

        assert_eq!(worksheet.row_infos().len(), 1);
        let row = &worksheet.row_infos()[0];
        assert_eq!(row.row, 3);
        assert_eq!(row.height, Some(24.5));
        assert!(row.hidden);
        assert_eq!(row.column_spans, vec![(2, 2)]);
    }

    #[test]
    fn auto_filter_range_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Filtered");
        sheet.set_cell(0, 0, "Header");
        sheet.set_auto_filter(0, 20, 0, 4);
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let auto_filter = reader.worksheet(0).unwrap().auto_filter().unwrap();
        assert_eq!(
            auto_filter,
            crate::xlsb::XlsbAutoFilter {
                first_row: 0,
                last_row: 20,
                first_column: 0,
                last_column: 4,
            }
        );
    }

    #[test]
    fn classic_and_extension_validations_survive_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Validated");
        sheet.set_cell(0, 0, 5);

        let mut classic = DataValidation::new(1, "A1:A10 C1:C10".to_string());
        classic.operator = 0;
        classic.formula1 = Some("1".to_string());
        classic.formula2 = Some("10".to_string());
        classic.ime_mode = 4;
        classic.show_input_message = true;
        classic.input_title = Some("Number".to_string());
        classic.input_text = Some("Enter 1 through 10".to_string());
        sheet.add_data_validation(classic);

        let mut extension = DataValidation::new(7, "B1:B20".to_string());
        extension.formula1 = Some("Source!A1>0".to_string());
        extension.record_kind = DataValidationRecordKind::Extension14;
        sheet.add_data_validation(extension);
        sheet.set_data_validation_settings(DataValidationSettings {
            input_prompts_disabled: true,
            prompt_x: 120,
            prompt_y: 240,
        });
        sheet.set_data_validation14_settings(DataValidationSettings {
            input_prompts_disabled: false,
            prompt_x: 12,
            prompt_y: 24,
        });
        workbook.add_worksheet(sheet);
        let mut source = MutableXlsbWorksheet::new("Source");
        source.set_cell(0, 0, 1);
        workbook.add_worksheet(source);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();
        assert_eq!(worksheet.data_validations().len(), 2);
        assert_eq!(
            worksheet.data_validations()[0].record_kind,
            DataValidationRecordKind::Classic
        );
        assert_eq!(worksheet.data_validations()[0].cell_ranges, "A1:A10 C1:C10");
        assert_eq!(
            worksheet.data_validations()[0].formula1.as_deref(),
            Some("1")
        );
        assert_eq!(
            worksheet.data_validations()[0].formula2.as_deref(),
            Some("10")
        );
        assert_eq!(worksheet.data_validations()[0].ime_mode, 4);
        assert_eq!(
            worksheet.data_validations()[1].record_kind,
            DataValidationRecordKind::Extension14
        );
        assert_eq!(
            worksheet.data_validations()[1].formula1.as_deref(),
            Some("(Source!A1>0)")
        );
        assert_eq!(
            worksheet.data_validation_settings(),
            Some(DataValidationSettings {
                input_prompts_disabled: true,
                prompt_x: 120,
                prompt_y: 240,
            })
        );
        assert_eq!(
            worksheet.data_validation14_settings(),
            Some(DataValidationSettings {
                input_prompts_disabled: false,
                prompt_x: 12,
                prompt_y: 24,
            })
        );
    }

    #[test]
    fn classic_conditional_formatting_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Formatted");
        sheet.set_cell(0, 0, 5);
        let mut formatting = ConditionalFormatting::new(vec!["A1:A10 C1:C10".to_string()]);
        formatting.pivot_only = true;

        let mut expression = ConditionalFormattingRule::new(CfRuleType::Expression, 1);
        expression.formula_texts.push("Source!A1>0".to_string());
        expression.stop_if_true = true;
        formatting.add_rule(expression);

        let mut scale = ConditionalFormattingRule::new(CfRuleType::ColorScale, 2);
        scale.color_scale = Some(ColorScale::new(
            Cfvo::new(2, None),
            Cfvo::new(7, Some("Source!A1".to_string())),
            0xffff_0000,
            0xff00_ff00,
        ));
        formatting.add_rule(scale);

        let mut bar = ConditionalFormattingRule::new(CfRuleType::DataBar, 3);
        bar.data_bar = Some(DataBar::new(
            Cfvo::new(2, None),
            Cfvo::new(3, None),
            0xff44_72c4,
        ));
        formatting.add_rule(bar);

        let mut icons = ConditionalFormattingRule::new(CfRuleType::IconSet, 4);
        icons.icon_set = Some(IconSet::new(
            0,
            vec![
                Cfvo::new(1, Some("0".to_string())),
                Cfvo::new(4, Some("33".to_string())),
                Cfvo::new(4, Some("67".to_string())),
            ],
        ));
        formatting.add_rule(icons);
        sheet.add_conditional_formatting(formatting);
        workbook.add_worksheet(sheet);

        let mut source = MutableXlsbWorksheet::new("Source");
        source.set_cell(0, 0, 10);
        workbook.add_worksheet(source);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();
        let formatting = &worksheet.conditional_formattings()[0];
        assert_eq!(formatting.ranges, ["A1:A10", "C1:C10"]);
        assert!(formatting.pivot_only);
        assert_eq!(formatting.rules.len(), 4);
        assert_eq!(formatting.rules[0].formula_texts, ["(Source!A1>0)"]);
        assert!(formatting.rules[0].stop_if_true);
        let scale = formatting.rules[1].color_scale.as_ref().unwrap();
        assert_eq!(scale.max_cfvo.value.as_deref(), Some("Source!A1"));
        assert_eq!(scale.min_color, 0xffff_0000);
        assert_eq!(scale.max_color, 0xff00_ff00);
        assert!(formatting.rules[2].data_bar.is_some());
        let icons = formatting.rules[3].icon_set.as_ref().unwrap();
        assert_eq!(icons.icon_set_type, 0);
        assert_eq!(icons.cfvos.len(), 3);
    }

    #[test]
    fn extension_conditional_formatting_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Formatted");
        sheet.set_cell(0, 0, 5);
        let mut formatting = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        formatting.record_kind = ConditionalFormattingRecordKind::Extension14;

        let mut rule = ConditionalFormattingRule::new(CfRuleType::DataBar, 1);
        rule.extension14 = Some(ConditionalFormattingRule14Metadata {
            priority: 1,
            unused: 0xCAFE_BABE,
            guid: [0x2a; 16],
            guid_present: true,
            linked_classic_priority: None,
        });
        let mut maximum = Cfvo::new(7, Some("Source!A1".to_string()));
        maximum.greater_than_or_equal = false;
        let mut bar = DataBar14::new(
            Cfvo::new(8, None),
            maximum,
            ConditionalFormatColor::from_argb(0xff44_72c4),
        );
        bar.min_length = 4;
        bar.max_length = 96;
        bar.gradient = false;
        rule.data_bar14 = Some(bar);
        formatting.add_rule(rule);
        sheet.add_conditional_formatting(formatting);
        workbook.add_worksheet(sheet);

        let mut source = MutableXlsbWorksheet::new("Source");
        source.set_cell(0, 0, 10);
        workbook.add_worksheet(source);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();
        let formatting = &worksheet.conditional_formattings()[0];
        assert_eq!(
            formatting.record_kind,
            ConditionalFormattingRecordKind::Extension14
        );
        let rule = &formatting.rules[0];
        assert_eq!(rule.extension14.unwrap().unused, 0xCAFE_BABE);
        let bar = rule.data_bar14.as_ref().unwrap();
        assert_eq!(bar.min_cfvo.cfvo_type, 8);
        assert_eq!(bar.max_cfvo.value.as_deref(), Some("Source!A1"));
        assert!(!bar.max_cfvo.greater_than_or_equal);
        assert_eq!((bar.min_length, bar.max_length), (4, 96));
        assert!(!bar.gradient);
        assert_eq!(bar.positive_color.unwrap().argb, Some(0xff44_72c4));
    }

    #[test]
    fn extended_data_bar_resolves_its_classic_rule_guid() {
        let guid = [0x7b; 16];
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Formatted");
        sheet.set_cell(0, 0, -5);

        let mut classic = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        let mut classic_rule = ConditionalFormattingRule::new(CfRuleType::DataBar, 1);
        classic_rule.classic_extension_guid = Some(guid);
        classic_rule.data_bar = Some(DataBar::new(
            Cfvo::new(2, None),
            Cfvo::new(3, None),
            0xff44_72c4,
        ));
        classic.add_rule(classic_rule);
        sheet.add_conditional_formatting(classic);

        let mut extension = ConditionalFormatting::new(vec!["A1:A10".to_string()]);
        extension.record_kind = ConditionalFormattingRecordKind::Extension14;
        let mut extension_rule = ConditionalFormattingRule::new(CfRuleType::DataBar, 0);
        extension_rule.template = 0;
        extension_rule.extension14 = Some(ConditionalFormattingRule14Metadata {
            priority: -1,
            unused: 0,
            guid,
            guid_present: true,
            linked_classic_priority: Some(1),
        });
        let mut bar = DataBar14::new(
            Cfvo::new(8, None),
            Cfvo::new(9, None),
            ConditionalFormatColor::from_argb(0xff44_72c4),
        );
        bar.min_length = 0;
        bar.max_length = 100;
        bar.positive_color = None;
        bar.negative_color = Some(ConditionalFormatColor::from_argb(0xffff_0000));
        bar.custom_negative_fill = true;
        extension_rule.data_bar14 = Some(bar);
        extension.add_rule(extension_rule);
        sheet.add_conditional_formatting(extension);
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet(0).unwrap();
        assert_eq!(worksheet.conditional_formattings().len(), 2);
        assert_eq!(
            worksheet.conditional_formattings()[0].rules[0].classic_extension_guid,
            Some(guid)
        );
        let extension_rule = &worksheet.conditional_formattings()[1].rules[0];
        assert_eq!(
            extension_rule.extension14.unwrap().linked_classic_priority,
            Some(1)
        );
        assert_eq!(
            extension_rule
                .data_bar14
                .as_ref()
                .unwrap()
                .negative_color
                .unwrap()
                .argb,
            Some(0xffff_0000)
        );
    }

    #[test]
    fn sheet_protection_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Protected");
        sheet.set_cell(0, 0, "locked");
        sheet.set_sheet_protection(Some(SheetProtection {
            password_hash: Some(0x5A3C),
            objects: Some(true),
            format_cells: Some(false),
            sort: Some(false),
            ..SheetProtection::default()
        }));
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let protection = reader.worksheet(0).unwrap().sheet_protection().unwrap();
        assert_eq!(protection.password_hash, Some(0x5A3C));
        assert!(protection.locked);
        assert!(!protection.allow_edit_objects);
        assert!(protection.allow_edit_scenarios);
        assert!(protection.allow_format_cells);
        assert!(!protection.allow_format_columns);
        assert!(protection.allow_sort);
        assert!(protection.allow_select_locked_cells);
    }

    #[test]
    fn test_workbook_writer_default() {
        let workbook: XlsbWorkbookWriter = Default::default();
        assert_eq!(workbook.worksheet_count(), 0);
        assert!(!workbook.is_1904);
    }

    #[test]
    fn test_get_worksheet_mut() {
        let mut workbook = XlsbWorkbookWriter::new();
        let sheet = MutableXlsbWorksheet::new("Sheet1");
        workbook.add_worksheet(sheet);

        let sheet_ref = workbook.get_worksheet_mut(0);
        assert!(sheet_ref.is_some());
        assert_eq!(sheet_ref.unwrap().name(), "Sheet1");

        assert!(workbook.get_worksheet_mut(99).is_none());
    }

    #[test]
    fn test_styles_accessor() {
        let workbook = XlsbWorkbookWriter::new();
        let styles = workbook.styles();
        // Just verify it returns a reference
        let _ = styles;
    }

    #[test]
    fn test_styles_mut_accessor() {
        let mut workbook = XlsbWorkbookWriter::new();
        let styles = workbook.styles_mut();
        // Just verify it returns a mutable reference
        let _ = styles;
    }

    #[test]
    fn test_add_multiple_worksheets() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet2"));
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet3"));

        assert_eq!(workbook.worksheet_count(), 3);
    }

    #[test]
    fn test_create_app_xml() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet2"));

        let app_xml = workbook.create_app_xml();
        assert!(app_xml.contains("<Application>The Litchi Rust Library</Application>"));
        assert!(app_xml.contains("<vt:i4>2</vt:i4>")); // Sheet count
        assert!(app_xml.contains("<vt:lpstr>Sheet1</vt:lpstr>"));
        assert!(app_xml.contains("<vt:lpstr>Sheet2</vt:lpstr>"));
    }

    #[test]
    fn test_create_core_xml() {
        let workbook = XlsbWorkbookWriter::new();
        let core_xml = workbook.create_core_xml();

        assert!(core_xml.contains("<dc:creator>The Litchi Rust Library</dc:creator>"));
        assert!(
            core_xml.contains("<cp:lastModifiedBy>The Litchi Rust Library</cp:lastModifiedBy>")
        );
        assert!(core_xml.contains("<cp:coreProperties"));
        assert!(core_xml.contains("</cp:coreProperties>"));
    }

    #[test]
    fn test_create_minimal_theme() {
        let workbook = XlsbWorkbookWriter::new();
        let theme = workbook.create_minimal_theme();

        assert!(theme.contains("<a:theme"));
        assert!(theme.contains("</a:theme>"));
    }

    #[test]
    fn test_format_w3cdtf() {
        let timestamp = format_w3cdtf(0); // Unix epoch
        assert!(timestamp.contains("T"));
        assert!(timestamp.ends_with("Z"));
        assert!(timestamp.starts_with("1970-"));
    }

    #[test]
    fn test_add_named_range() {
        use crate::xlsb::named_ranges::NamedRange;

        let mut workbook = XlsbWorkbookWriter::new();
        let named_range = NamedRange::new("TestRange".to_string(), None).with_formula(vec![
            crate::xlsb::formula::ptg_types::PTG_INT,
            1,
            0,
        ]);
        workbook.add_named_range(named_range);
        assert_eq!(workbook.named_ranges.len(), 1);
        assert_eq!(workbook.named_ranges[0].name, "TestRange");
    }

    #[test]
    fn defined_name_survives_package_roundtrip() {
        use crate::xlsb::named_ranges::{NamedRange, create_area3d_formula};

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data Sheet"));
        workbook.add_named_range(
            NamedRange::new("SalesData".to_string(), None)
                .with_formula(create_area3d_formula(0, 1, 3, 1, 1).unwrap()),
        );
        let mut summary = MutableXlsbWorksheet::new("Summary");
        summary.set_cell(
            0,
            0,
            CellValue::Formula {
                formula: "SalesData".to_string(),
                cached_value: Some(Box::new(CellValue::Float(0.0))),
                is_array: false,
                array_range: None,
            },
        );
        summary.set_cell(
            0,
            1,
            CellValue::Formula {
                formula: "'Data Sheet'!$B$2".to_string(),
                cached_value: Some(Box::new(CellValue::Float(0.0))),
                is_array: false,
                array_range: None,
            },
        );
        workbook.add_worksheet(summary);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        assert_eq!(reader.defined_names(), &["SalesData"]);
        let summary = reader.worksheet_by_index(1).unwrap();
        assert!(matches!(
            summary.cell_value(0, 0).unwrap().as_ref(),
            CellValue::Formula { formula, .. } if formula == "SalesData"
        ));
        assert!(matches!(
            summary.cell_value(0, 1).unwrap().as_ref(),
            CellValue::Formula { formula, .. } if formula == "'Data Sheet'!$B$2"
        ));
    }

    #[test]
    fn sheet_range_formula_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data Sheet"));
        workbook.add_worksheet(MutableXlsbWorksheet::new("Middle"));
        let mut summary = MutableXlsbWorksheet::new("Summary");
        for col in 0..2 {
            summary.set_cell(
                0,
                col,
                CellValue::Formula {
                    formula: "SUM('Data Sheet:Summary'!A1)".to_string(),
                    cached_value: Some(Box::new(CellValue::Float(0.0))),
                    is_array: false,
                    array_range: None,
                },
            );
        }
        workbook.add_worksheet(summary);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let summary = reader.worksheet_by_index(2).unwrap();
        for col in 0..2 {
            assert!(matches!(
                summary.cell_value(0, col).unwrap().as_ref(),
                CellValue::Formula { formula, .. }
                    if formula == "SUM('Data Sheet:Summary'!A1)"
            ));
        }
    }

    #[test]
    fn contextual_grouped_formulas_survive_package_roundtrip() {
        use crate::xlsb::named_ranges::{NamedRange, create_area3d_formula};

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data"));
        workbook.add_worksheet(MutableXlsbWorksheet::new("Middle"));
        workbook.add_named_range(
            NamedRange::new("Rate".to_string(), None)
                .with_formula(create_area3d_formula(0, 0, 0, 0, 0).unwrap()),
        );
        let mut summary = MutableXlsbWorksheet::new("Summary");
        summary
            .set_array_formula(0, 0, 0, 1, "SUM('Data:Middle'!A1)+Rate")
            .unwrap();
        summary
            .set_shared_formula(0, 2, 1, 2, "Data!A1+$A1")
            .unwrap();
        summary.set_cell(
            0,
            3,
            CellValue::Formula {
                formula: "Middle!A1".to_string(),
                cached_value: None,
                is_array: true,
                array_range: Some("D1:E1".to_string()),
            },
        );
        workbook.add_worksheet(summary);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let summary = reader.worksheet_by_index(2).unwrap();
        for col in 0..=1 {
            assert!(matches!(
                summary.cell_value(0, col).unwrap().as_ref(),
                CellValue::Formula {
                    formula,
                    is_array: true,
                    array_range: Some(range),
                    ..
                } if formula == "(SUM('Data:Middle'!A1)+Rate)" && range == "A1:B1"
            ));
        }
        assert!(matches!(
            summary.cell_value(0, 2).unwrap().as_ref(),
            CellValue::Formula { formula, is_array: false, .. }
                if formula == "(Data!A1+$A1)"
        ));
        assert!(matches!(
            summary.cell_value(1, 2).unwrap().as_ref(),
            CellValue::Formula { formula, is_array: false, .. }
                if formula == "(Data!A1+$A2)"
        ));
        for col in 3..=4 {
            assert!(matches!(
                summary.cell_value(0, col).unwrap().as_ref(),
                CellValue::Formula {
                    formula,
                    is_array: true,
                    array_range: Some(range),
                    ..
                } if formula == "Middle!A1" && range == "D1:E1"
            ));
        }

        workbook.get_worksheet_mut(0).unwrap().set_name("Renamed");
        assert!(workbook.save(Cursor::new(Vec::new())).is_err());

        let mut invalid = MutableXlsbWorksheet::new("Invalid");
        assert!(
            invalid
                .set_array_formula(0, 0, 0, 0, "NOT_A_REAL_FUNCTION(1)")
                .is_err()
        );
    }

    #[test]
    fn rejects_ambiguous_formula_metadata_before_writing() {
        use crate::xlsb::named_ranges::{NamedRange, create_area3d_formula};

        let mut duplicate_sheets = XlsbWorkbookWriter::new();
        duplicate_sheets.add_worksheet(MutableXlsbWorksheet::new("Data"));
        duplicate_sheets.add_worksheet(MutableXlsbWorksheet::new("data"));
        assert!(duplicate_sheets.save(Cursor::new(Vec::new())).is_err());

        let mut invalid_sheet = XlsbWorkbookWriter::new();
        invalid_sheet.add_worksheet(MutableXlsbWorksheet::new("Data/2026"));
        assert!(invalid_sheet.save(Cursor::new(Vec::new())).is_err());

        let mut duplicate_names = XlsbWorkbookWriter::new();
        duplicate_names.add_worksheet(MutableXlsbWorksheet::new("Data"));
        duplicate_names.add_named_range(
            NamedRange::new("Rate".to_string(), None)
                .with_formula(create_area3d_formula(0, 0, 0, 0, 0).unwrap()),
        );
        duplicate_names.add_named_range(
            NamedRange::new("rate".to_string(), None)
                .with_formula(create_area3d_formula(0, 1, 1, 0, 0).unwrap()),
        );
        assert!(duplicate_names.save(Cursor::new(Vec::new())).is_err());
    }

    #[test]
    fn contextual_formula_tokens_are_not_cached_across_saves() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(MutableXlsbWorksheet::new("Data"));
        let mut summary = MutableXlsbWorksheet::new("Summary");
        summary.set_cell(
            0,
            0,
            CellValue::Formula {
                formula: "Data!A1".to_string(),
                cached_value: None,
                is_array: false,
                array_range: None,
            },
        );
        workbook.add_worksheet(summary);
        workbook.save(Cursor::new(Vec::new())).unwrap();

        workbook.get_worksheet_mut(0).unwrap().set_name("Renamed");
        assert!(workbook.save(Cursor::new(Vec::new())).is_err());
    }

    #[test]
    fn formula_survives_workbook_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Calculations");
        sheet.set_cell(0, 0, 2.0);
        sheet.set_cell(0, 1, 3.0);
        sheet.set_cell(
            0,
            2,
            CellValue::Formula {
                formula: "A1+B1".to_string(),
                cached_value: Some(Box::new(CellValue::Float(5.0))),
                is_array: false,
                array_range: None,
            },
        );
        sheet.set_cell(
            1,
            0,
            CellValue::Formula {
                formula: "\"result\"".to_string(),
                cached_value: Some(Box::new(CellValue::String("result".to_string()))),
                is_array: false,
                array_range: None,
            },
        );
        sheet.set_cell(
            1,
            1,
            CellValue::Formula {
                formula: "1=1".to_string(),
                cached_value: Some(Box::new(CellValue::Bool(true))),
                is_array: false,
                array_range: None,
            },
        );
        sheet.set_cell(
            1,
            2,
            CellValue::Formula {
                formula: "1/0".to_string(),
                cached_value: Some(Box::new(CellValue::Error("#DIV/0!".to_string()))),
                is_array: false,
                array_range: None,
            },
        );
        sheet.set_cell(
            2,
            0,
            CellValue::Formula {
                formula: "#REF!".to_string(),
                cached_value: Some(Box::new(CellValue::Error("#REF!".to_string()))),
                is_array: false,
                array_range: None,
            },
        );
        sheet.set_cell(
            2,
            1,
            CellValue::Formula {
                formula: "IF(TRUE,1,2)".to_string(),
                cached_value: Some(Box::new(CellValue::Float(1.0))),
                is_array: false,
                array_range: None,
            },
        );
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet_by_index(0).unwrap();
        let value = worksheet.cell_value(0, 2).unwrap();
        assert!(matches!(
            value.as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                ..
            } if formula == "(A1+B1)" && matches!(cached.as_ref(), CellValue::Float(5.0))
        ));
        assert!(matches!(
            worksheet.cell_value(1, 0).unwrap().as_ref(),
            CellValue::Formula {
                cached_value: Some(cached),
                ..
            } if matches!(cached.as_ref(), CellValue::String(value) if value == "result")
        ));
        assert!(matches!(
            worksheet.cell_value(1, 1).unwrap().as_ref(),
            CellValue::Formula {
                cached_value: Some(cached),
                ..
            } if matches!(cached.as_ref(), CellValue::Bool(true))
        ));
        assert!(matches!(
            worksheet.cell_value(1, 2).unwrap().as_ref(),
            CellValue::Formula {
                cached_value: Some(cached),
                ..
            } if matches!(cached.as_ref(), CellValue::Error(error) if error == "#DIV/0!")
        ));
        assert!(matches!(
            worksheet.cell_value(2, 0).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                ..
            } if formula == "#REF!"
                && matches!(cached.as_ref(), CellValue::Error(error) if error == "#REF!")
        ));
        assert!(matches!(
            worksheet.cell_value(2, 1).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                ..
            } if formula == "IF(TRUE,1,2)"
                && matches!(cached.as_ref(), CellValue::Float(1.0))
        ));
    }

    #[test]
    fn array_and_shared_formulas_survive_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Grouped formulas");
        sheet.set_cell(2, 2, 10.0);
        sheet.set_cell(3, 2, 20.0);
        sheet.set_shared_formula(2, 2, 3, 2, "B3").unwrap();
        sheet.set_array_formula(0, 4, 1, 5, "A1*2").unwrap();
        // The core CellValue representation remains a supported compatibility
        // path; the writer fills missing follower records from array_range.
        sheet.set_cell(
            5,
            0,
            CellValue::Formula {
                formula: "1+1".to_string(),
                cached_value: None,
                is_array: true,
                array_range: Some("A6:B6".to_string()),
            },
        );
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet_by_index(0).unwrap();

        assert!(matches!(
            worksheet.cell_value(2, 2).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                is_array: false,
                ..
            } if formula == "B3" && matches!(cached.as_ref(), CellValue::Float(10.0))
        ));
        assert!(matches!(
            worksheet.cell_value(3, 2).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                is_array: false,
                ..
            } if formula == "B4" && matches!(cached.as_ref(), CellValue::Float(20.0))
        ));
        for row in 0..=1 {
            for col in 4..=5 {
                assert!(matches!(
                    worksheet.cell_value(row, col).unwrap().as_ref(),
                    CellValue::Formula {
                        formula,
                        is_array: true,
                        array_range: Some(range),
                        ..
                    } if formula == "(A1*2)" && range == "E1:F2"
                ));
            }
        }
        for col in 0..=1 {
            assert!(matches!(
                worksheet.cell_value(5, col).unwrap().as_ref(),
                CellValue::Formula {
                    formula,
                    is_array: true,
                    array_range: Some(range),
                    ..
                } if formula == "(1+1)" && range == "A6:B6"
            ));
        }
    }

    #[test]
    fn array_constant_survives_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let mut sheet = MutableXlsbWorksheet::new("Array constant");
        sheet.set_cell(
            0,
            0,
            CellValue::Formula {
                formula: "SUM({1,2;3,4})".to_string(),
                cached_value: Some(Box::new(CellValue::Float(10.0))),
                is_array: false,
                array_range: None,
            },
        );
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let worksheet = reader.worksheet_by_index(0).unwrap();
        assert!(matches!(
            worksheet.cell_value(0, 0).unwrap().as_ref(),
            CellValue::Formula {
                formula,
                cached_value: Some(cached),
                ..
            } if formula == "SUM({1,2;3,4})"
                && matches!(cached.as_ref(), CellValue::Float(10.0))
        ));
    }

    #[test]
    fn worksheet_charts_round_trip_through_binary_drawing_graphs() {
        use crate::xlsb::drawing::XlsbDrawingAnchorKind;
        use crate::xlsx::{ChartAnchor, WorksheetChart};
        use litchi_drawingml::chart::plot_area::TypeGroup;

        let bar = WorksheetChart::bar_chart_with_cache(
            "Quarterly sales",
            "Charts!$A$2:$A$4",
            &["Q1", "Q2", "Q3"],
            "Charts!$B$2:$B$4",
            &[10.0, 20.0, 30.0],
            ChartAnchor::with_offsets(1, 10, 1, 20, 8, 30, 15, 40),
        )
        .unwrap();
        let line = WorksheetChart::line_chart(
            "Trend",
            "Charts!$A$2:$A$4",
            "Charts!$B$2:$B$4",
            ChartAnchor::new(9, 1, 16, 15),
        )
        .unwrap();

        let mut sheet = MutableXlsbWorksheet::new("Charts");
        sheet.set_cell(0, 0, "Quarter");
        sheet.set_cell(0, 1, "Sales");
        sheet.add_chart(bar).unwrap();
        sheet.add_chart(line).unwrap();
        assert_eq!(sheet.charts().len(), 2);

        let pie = WorksheetChart::pie_chart(
            "Share",
            "Summary!$A$1:$A$3",
            "Summary!$B$1:$B$3",
            ChartAnchor::new(0, 0, 7, 12),
        )
        .unwrap();
        let mut summary = MutableXlsbWorksheet::new("Summary");
        summary.add_chart(pie).unwrap();

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(sheet);
        workbook.add_worksheet(summary);
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();

        let drawing = reader.sheet_drawing(0).expect("sheet drawing missing");
        assert_eq!(drawing.drawing.anchors.len(), 2);
        assert_eq!(drawing.charts.len(), 2);
        assert!(matches!(
            drawing.charts[0].chart.plot_area.type_groups.as_slice(),
            [TypeGroup::Bar(_)]
        ));
        assert!(matches!(
            drawing.charts[1].chart.plot_area.type_groups.as_slice(),
            [TypeGroup::Line(_)]
        ));
        let summary_drawing = reader
            .sheet_drawing(1)
            .expect("second sheet drawing missing");
        assert!(matches!(
            summary_drawing.charts[0]
                .chart
                .plot_area
                .type_groups
                .as_slice(),
            [TypeGroup::Pie(_)]
        ));
        match &drawing.drawing.anchors[0].anchor {
            XlsbDrawingAnchorKind::TwoCell { from, to, edit_as } => {
                assert_eq!((from.column, from.row), (1, 1));
                assert_eq!((from.column_offset, from.row_offset), (10, 20));
                assert_eq!((to.column, to.row), (8, 15));
                assert_eq!((to.column_offset, to.row_offset), (30, 40));
                assert!(edit_as.is_none());
            },
            other => panic!("unexpected chart anchor: {other:?}"),
        }

        let package = reader.opc_package();
        assert_eq!(
            package
                .iter_parts()
                .filter(|part| part.content_type() == ct::OFC_DRAWING)
                .count(),
            2
        );
        assert_eq!(
            package
                .iter_parts()
                .filter(|part| part.content_type() == ct::DML_CHART)
                .count(),
            3
        );
        let sheet_uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
        let sheet_part = package.get_part(&sheet_uri).unwrap();
        let drawing_record = litchi_xlsb::raw::Records::new(sheet_part.blob())
            .find_map(|record| {
                let record = record.unwrap();
                (record.kind() == kind::DRAWING).then_some(record)
            })
            .expect("BrtDrawing missing");
        let mut cursor = litchi_xlsb::raw::Cursor::new(drawing_record.payload(), "BrtDrawing");
        let drawing_rel_id = cursor.read_wide_string().unwrap();
        cursor.finish().unwrap();
        let relationship = sheet_part.rels().get(&drawing_rel_id).unwrap();
        assert_eq!(relationship.reltype(), rel::DRAWING);
        assert!(!relationship.is_external());
    }

    #[test]
    fn pivot_chart_round_trips_with_lossless_view_and_cache_graph() {
        use crate::xlsx::{ChartAnchor, WorksheetChart};

        let mut begin_view = vec![0u8; 32];
        begin_view[28..32].copy_from_slice(&1u32.to_le_bytes());
        let view_name = "RevenuePivot";
        begin_view.extend_from_slice(&(view_name.len() as u32).to_le_bytes());
        for unit in view_name.encode_utf16() {
            begin_view.extend_from_slice(&unit.to_le_bytes());
        }
        let mut view_bytes = Vec::new();
        {
            let mut writer = Writer::new(&mut view_bytes);
            writer
                .write_record(kind::BEGIN_SX_VIEW, &begin_view)
                .unwrap();
            writer
                .write_record(kind::BEGIN_SX_LOCATION, &[0; 36])
                .unwrap();
            writer.write_record(kind::END_SX_LOCATION, &[]).unwrap();
            writer.write_record(kind::END_SX_VIEW, &[]).unwrap();
        }
        let view = crate::xlsb::pivot_view::XlsbPivotTableViewPart::from_bytes(view_bytes.clone())
            .unwrap();

        let chart = WorksheetChart::line_chart(
            "Revenue",
            "Pivot Host!$A$2:$A$3",
            "Pivot Host!$B$2:$B$3",
            ChartAnchor::new(3, 0, 10, 14),
        )
        .unwrap()
        .into_pivot_chart(view_name)
        .unwrap();
        let mut sheet = MutableXlsbWorksheet::new("Pivot Host");
        sheet.add_pivot_table_view(view).unwrap();
        sheet.add_chart(chart).unwrap();

        let mut workbook = XlsbWorkbookWriter::new();
        let cache_id = workbook
            .add_pivot_cache(&crate::xlsb::pivot::PivotCacheDefinition::default())
            .unwrap();
        assert_eq!(cache_id, 1);
        workbook.add_worksheet(sheet);

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let bytes = output.into_inner();
        let package = OpcPackage::from_bytes(&bytes).unwrap();
        let sheet_part = package
            .get_part(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
            .unwrap();
        let view_relationship = sheet_part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rel::PIVOT_TABLE)
            .expect("worksheet PivotTable relationship missing");
        let view_part = package
            .get_part(&view_relationship.target_partname().unwrap())
            .unwrap();
        assert_eq!(view_part.blob(), view_bytes);
        let cache_relationship = view_part
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == rel::PIVOT_CACHE_DEFINITION)
            .expect("PivotTable cache relationship missing");
        assert_eq!(
            cache_relationship.target_partname().unwrap(),
            PackURI::new("/xl/pivotCache/pivotCacheDefinition1.bin").unwrap()
        );

        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(bytes)).unwrap();
        assert_eq!(reader.pivot_views().len(), 1);
        assert_eq!(reader.pivot_views()[0].name(), view_name);
        assert_eq!(reader.pivot_views()[0].cache_id(), 1);
        let drawing = reader
            .sheet_drawing(0)
            .expect("pivot chart drawing missing");
        let source = drawing.charts[0]
            .chart
            .pivot_source
            .as_ref()
            .expect("pivot source missing");
        assert_eq!(source.name, "'Pivot Host'!RevenuePivot");
    }

    #[test]
    fn pivot_chart_refuses_a_missing_view_binding() {
        use crate::xlsx::{ChartAnchor, WorksheetChart};

        let chart = WorksheetChart::line_chart(
            "Revenue",
            "Host!$A$1:$A$2",
            "Host!$B$1:$B$2",
            ChartAnchor::new(2, 0, 8, 12),
        )
        .unwrap()
        .into_pivot_chart("MissingPivot")
        .unwrap();
        let mut sheet = MutableXlsbWorksheet::new("Host");
        sheet.add_chart(chart).unwrap();
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(sheet);

        let error = workbook
            .save(Cursor::new(Vec::new()))
            .expect_err("missing PivotTable binding must fail");
        assert!(error.to_string().contains("missing pivot table"));
    }

    #[test]
    fn chart_resource_graphs_round_trip_for_worksheets_and_chart_sheets() {
        use crate::xlsx::{
            ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship,
            ChartRelationshipTarget, ChartUserShapesPart, ChartUserShapesRelationship,
            ChartUserShapesRelationshipTarget, WorksheetChart,
        };
        use litchi_drawingml::chart::{ChartExtensionList, ChartShapeProperties};

        let mut worksheet_chart = WorksheetChart::bar_chart(
            "Resources",
            "Data!$A$1:$A$2",
            "Data!$B$1:$B$2",
            ChartAnchor::new(1, 1, 8, 15),
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
                br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:example"><c:ext uri="resources"><x:reference r:id="rId1" r:link="rId10"/></c:ext></c:extLst>"#.to_vec(),
            )
            .unwrap(),
        );
        worksheet_chart = worksheet_chart
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId9".to_string(),
                relationship_type: rel::IMAGE.to_string(),
                target: ChartRelationshipTarget::Embedded {
                    data: b"chart background".to_vec(),
                    content_type: "image/png".to_string(),
                    extension: "png".to_string(),
                },
            })
            .with_additional_relationship(ChartRelationship {
                relationship_id: "rId10".to_string(),
                relationship_type: rel::HYPERLINK.to_string(),
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
                    relationship_type: rel::IMAGE.to_string(),
                    target: ChartUserShapesRelationshipTarget::Embedded {
                        data: b"shape image".to_vec(),
                        content_type: "image/png".to_string(),
                        extension: "png".to_string(),
                    },
                }],
            });

        let chart_sheet_chart = WorksheetChart::line_chart(
            "Linked",
            "Data!$A$1:$A$2",
            "Data!$B$1:$B$2",
            ChartAnchor::new(0, 0, 5, 10),
        )
        .unwrap()
        .with_external_data_part(
            ChartExternalDataPart::linked_package("https://example.test/data.xlsx"),
            Some(true),
        );

        let mut workbook = XlsbWorkbookWriter::new();
        let mut data = MutableXlsbWorksheet::new("Data");
        data.add_chart(worksheet_chart).unwrap();
        workbook.add_worksheet(data);
        workbook
            .add_chart_sheet(MutableXlsbChartSheet::new(
                "Linked Chart",
                chart_sheet_chart,
            ))
            .unwrap();

        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();

        let worksheet_chart = &reader.sheet_drawing(0).unwrap().charts[0];
        match &worksheet_chart.external_data_part.as_ref().unwrap().target {
            ChartExternalDataTarget::Embedded { data, .. } => {
                assert_eq!(data, b"PK chart workbook");
            },
            other => panic!("unexpected worksheet chart external data: {other:?}"),
        }
        let user_shapes = worksheet_chart.user_shapes_part.as_ref().unwrap();
        assert_eq!(user_shapes.relationships.len(), 1);
        match &user_shapes.relationships[0].target {
            ChartRelationshipTarget::Embedded { data, .. } => {
                assert_eq!(data, b"shape image");
            },
            other => panic!("unexpected user-shapes target: {other:?}"),
        }
        assert_eq!(worksheet_chart.additional_relationships.len(), 2);
        let background = worksheet_chart
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId9")
            .unwrap();
        match &background.target {
            ChartRelationshipTarget::Embedded { data, .. } => {
                assert_eq!(data, b"chart background");
            },
            other => panic!("unexpected background target: {other:?}"),
        }
        let hyperlink = worksheet_chart
            .additional_relationships
            .iter()
            .find(|relationship| relationship.relationship_id == "rId10")
            .unwrap();
        match &hyperlink.target {
            ChartRelationshipTarget::External { target } => {
                assert_eq!(target, "https://example.test/chart");
            },
            other => panic!("unexpected hyperlink target: {other:?}"),
        }

        let chart_sheet_chart = &reader.sheet_drawing(1).unwrap().charts[0];
        match &chart_sheet_chart
            .external_data_part
            .as_ref()
            .unwrap()
            .target
        {
            ChartExternalDataTarget::Linked { target } => {
                assert_eq!(target, "https://example.test/data.xlsx");
            },
            other => panic!("unexpected chart-sheet external data: {other:?}"),
        }
        assert!(chart_sheet_chart.user_shapes_part.is_none());
        assert!(chart_sheet_chart.additional_relationships.is_empty());
    }

    #[test]
    fn worksheet_chart_validation_and_crud_are_lossless_or_refuse() {
        use crate::xlsx::{
            ChartAnchor, ChartRelationship, ChartRelationshipTarget, ChartUserShapesPart,
            WorksheetChart,
        };

        let mut sheet = MutableXlsbWorksheet::new("Charts");
        let valid = WorksheetChart::bar_chart(
            "Valid",
            "Charts!$A$1:$A$2",
            "Charts!$B$1:$B$2",
            ChartAnchor::new(1, 1, 8, 15),
        )
        .unwrap();
        sheet.add_chart(valid.clone()).unwrap();

        let mut descending = valid.clone();
        descending.anchor.to_col = 0;
        assert!(sheet.add_chart(descending).is_err());
        assert_eq!(sheet.charts().len(), 1);

        let mut mismatched_external_data = valid.clone();
        mismatched_external_data.chart.external_data =
            Some(litchi_drawingml::chart::ChartExternalData::pending());
        assert!(sheet.add_chart(mismatched_external_data).is_err());
        assert_eq!(sheet.charts().len(), 1);

        let invalid_user_shapes = valid.clone().with_user_shapes_part(ChartUserShapesPart::new(
            br#"<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="rId5"/></c:userShapes>"#.to_vec(),
        ));
        assert!(sheet.add_chart(invalid_user_shapes).is_err());
        let invalid_relationship = valid
            .clone()
            .with_additional_relationship(ChartRelationship {
                relationship_id: "not an id".to_string(),
                relationship_type: rel::HYPERLINK.to_string(),
                target: ChartRelationshipTarget::External {
                    target: "https://example.test".to_string(),
                },
            });
        assert!(sheet.add_chart(invalid_relationship).is_err());
        assert_eq!(sheet.charts().len(), 1);

        let removed = sheet.remove_chart(0).unwrap();
        assert_eq!(removed.anchor.from_col, 1);
        assert!(sheet.charts().is_empty());
        assert!(sheet.remove_chart(0).is_err());
        sheet.add_chart(valid).unwrap();
        sheet.clear_charts();
        assert!(sheet.charts().is_empty());
    }

    #[test]
    fn worksheet_images_round_trip_with_charts_in_one_drawing_graph() {
        use crate::xlsb::{
            XlsbDrawingAnchorKind, XlsbDrawingObject, XlsbWorksheetImage, XlsbWorksheetImageFormat,
        };
        use crate::xlsx::{
            ChartAnchor, WorksheetChart, XlsxDrawingObject, XlsxEmu, XlsxEmuExtent, XlsxEmuOffset,
            XlsxShapeAnchor, XlsxShapePreset,
        };

        const PNG_1X1: &[u8] = &[
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00,
            0x00, 0x1F, 0x15, 0xC4, 0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0xF0, 0x1F, 0x00, 0x05, 0x00, 0x01, 0xFF, 0x89, 0x99,
            0x03, 0x5D, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ];
        const SVG: &[u8] =
            br#"<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"><path d="M0 0h1v1z"/></svg>"#;
        const GIF_1X1: &[u8] = &[
            0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00,
            0x00, 0x00, 0xFF, 0xFF, 0xFF, 0x21, 0xF9, 0x04, 0x01, 0x00, 0x00, 0x00, 0x00, 0x2C,
            0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44, 0x01, 0x00,
            0x3B,
        ];

        let png_anchor = ChartAnchor::with_offsets(1, 10, 2, 20, 5, 30, 8, 40);
        let png =
            XlsbWorksheetImage::new(PNG_1X1.to_vec(), XlsbWorksheetImageFormat::Png, png_anchor)
                .unwrap()
                .with_description("Logo & <mark>")
                .unwrap();
        let svg = XlsbWorksheetImage::new(
            SVG.to_vec(),
            XlsbWorksheetImageFormat::Svg,
            ChartAnchor::new(6, 2, 9, 8),
        )
        .unwrap();
        let chart = WorksheetChart::line_chart(
            "Trend",
            "Pictures!$A$1:$A$2",
            "Pictures!$B$1:$B$2",
            ChartAnchor::new(10, 2, 17, 16),
        )
        .unwrap();

        let mut sheet = MutableXlsbWorksheet::new("Pictures");
        sheet.add_image(png).unwrap();
        sheet.add_image(svg).unwrap();
        sheet.add_chart(chart).unwrap();
        sheet
            .add_text_box(
                "Caption",
                XlsxShapeAnchor::Absolute {
                    position: XlsxEmuOffset {
                        x: XlsxEmu(100_000),
                        y: XlsxEmu(100_000),
                    },
                    extent: XlsxEmuExtent {
                        width: XlsxEmu(1_000_000),
                        height: XlsxEmu(500_000),
                    },
                },
                XlsxShapePreset::Rectangle,
                "Mixed drawing",
            )
            .unwrap();
        let mut image_only = MutableXlsbWorksheet::new("Image only");
        image_only
            .add_image(
                XlsbWorksheetImage::new(
                    GIF_1X1.to_vec(),
                    XlsbWorksheetImageFormat::Gif,
                    ChartAnchor::new(0, 0, 2, 3),
                )
                .unwrap(),
            )
            .unwrap();
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(sheet);
        workbook.add_worksheet(image_only);
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();

        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let drawing = reader.sheet_drawing(0).expect("sheet drawing missing");
        assert_eq!(drawing.images.len(), 2);
        assert_eq!(drawing.charts.len(), 1);
        assert_eq!(drawing.images[0].format, XlsbWorksheetImageFormat::Png);
        assert_eq!(drawing.images[0].data.as_ref(), PNG_1X1);
        assert_eq!(
            drawing.images[0].description.as_deref(),
            Some("Logo & <mark>")
        );
        assert_eq!(drawing.images[1].format, XlsbWorksheetImageFormat::Svg);
        assert_eq!(drawing.images[1].data.as_ref(), SVG);
        assert_eq!(drawing.images[0].rel_id, "rId1");
        assert_eq!(drawing.images[1].rel_id, "rId2");
        assert_eq!(drawing.charts[0].rel_id, "rId3");
        assert_eq!(drawing.drawing.anchors.len(), 4);
        assert_eq!(drawing.shapes.len(), 1);
        let XlsxDrawingObject::Shape(caption) = &drawing.shapes[0].object else {
            panic!("expected mixed-drawing caption");
        };
        assert_eq!(caption.non_visual.id, Some(4));
        match &drawing.drawing.anchors[0].anchor {
            XlsbDrawingAnchorKind::TwoCell { from, to, .. } => {
                assert_eq!(
                    (from.column, from.column_offset, from.row, from.row_offset),
                    (1, 10, 2, 20)
                );
                assert_eq!(
                    (to.column, to.column_offset, to.row, to.row_offset),
                    (5, 30, 8, 40)
                );
            },
            other => panic!("unexpected image anchor: {other:?}"),
        }
        assert!(matches!(
            &drawing.drawing.anchors[0].object,
            XlsbDrawingObject::Picture {
                embed_rel_id: Some(rel_id),
                ..
            } if rel_id == "rId1"
        ));
        assert!(matches!(
            &drawing.drawing.anchors[2].object,
            XlsbDrawingObject::GraphicFrame(frame)
                if frame.rel_id.as_deref() == Some("rId3")
        ));
        let second_drawing = reader
            .sheet_drawing(1)
            .expect("image-only sheet drawing missing");
        assert_eq!(second_drawing.images.len(), 1);
        assert!(second_drawing.charts.is_empty());
        assert_eq!(
            second_drawing.images[0].format,
            XlsbWorksheetImageFormat::Gif
        );
        assert_eq!(second_drawing.images[0].data.as_ref(), GIF_1X1);

        let package = reader.opc_package();
        assert_eq!(
            package
                .iter_parts()
                .filter(|part| matches!(part.content_type(), ct::PNG | ct::GIF | "image/svg+xml"))
                .count(),
            3
        );
        for part_name in [
            "/xl/media/image1.png",
            "/xl/media/image2.svg",
            "/xl/media/image3.gif",
        ] {
            assert!(package.get_part(&PackURI::new(part_name).unwrap()).is_ok());
        }
    }

    #[test]
    fn worksheet_image_validation_and_crud_are_lossless_or_refuse() {
        use crate::xlsb::{XlsbWorksheetImage, XlsbWorksheetImageFormat};
        use crate::xlsx::ChartAnchor;

        assert!(
            XlsbWorksheetImage::new(
                b"not a png".to_vec(),
                XlsbWorksheetImageFormat::Png,
                ChartAnchor::new(0, 0, 1, 1),
            )
            .is_err()
        );
        assert!(
            XlsbWorksheetImage::new(
                b"<not-svg/>".to_vec(),
                XlsbWorksheetImageFormat::Svg,
                ChartAnchor::new(0, 0, 1, 1),
            )
            .is_err()
        );
        assert!(
            XlsbWorksheetImage::new(
                b"GIF89a".to_vec(),
                XlsbWorksheetImageFormat::Gif,
                ChartAnchor::new(2, 2, 1, 1),
            )
            .is_err()
        );

        let valid = XlsbWorksheetImage::new(
            b"GIF89a".to_vec(),
            XlsbWorksheetImageFormat::Gif,
            ChartAnchor::new(0, 0, 1, 1),
        )
        .unwrap();
        let mut sheet = MutableXlsbWorksheet::new("Pictures");
        sheet.add_image(valid.clone()).unwrap();
        assert_eq!(sheet.images().len(), 1);
        assert!(
            valid
                .clone()
                .with_description("invalid\u{0}description")
                .is_err()
        );
        assert_eq!(sheet.images().len(), 1);
        let removed = sheet.remove_image(0).unwrap();
        assert_eq!(removed.format(), XlsbWorksheetImageFormat::Gif);
        assert!(sheet.images().is_empty());
        assert!(sheet.remove_image(0).is_err());
        sheet.add_image(valid).unwrap();
        sheet.clear_images();
        assert!(sheet.images().is_empty());
    }

    #[test]
    fn worksheet_shapes_groups_and_connectors_round_trip() {
        use crate::xlsx::writer::{
            XlsxConnectionEndSpec, XlsxConnectionShapeSpec, XlsxGroupSpec, XlsxShapeSpec,
        };
        use crate::xlsx::{
            XlsxCellMarker, XlsxDrawingObject, XlsxEditAs, XlsxEmu, XlsxEmuExtent, XlsxEmuOffset,
            XlsxGroupTransform, XlsxShapeAnchor, XlsxShapePreset,
        };

        fn marker(column: u32, row: u32) -> XlsxCellMarker {
            XlsxCellMarker {
                column,
                row,
                column_offset: XlsxEmu(0),
                row_offset: XlsxEmu(0),
            }
        }

        let two_cell = XlsxShapeAnchor::TwoCell {
            from: marker(0, 0),
            to: marker(3, 4),
            edit_as: XlsxEditAs::OneCell,
        };
        let child_anchor = XlsxShapeAnchor::TwoCell {
            from: marker(0, 0),
            to: marker(1, 1),
            edit_as: XlsxEditAs::TwoCell,
        };
        let mut standalone = XlsxShapeSpec::text_box(
            "Standalone",
            two_cell,
            XlsxShapePreset::RoundRectangle,
            "A\nB",
        );
        standalone.description = Some("Typed XLSB text box".to_string());
        standalone.paragraphs[0].runs[0].bold = Some(true);
        standalone.paragraphs[0].runs[0].font_size_hundredths = Some(1_400);

        let group_anchor = XlsxShapeAnchor::OneCell {
            from: marker(4, 1),
            extent: XlsxEmuExtent {
                width: XlsxEmu(4_000_000),
                height: XlsxEmu(2_000_000),
            },
        };
        let mut group = XlsxGroupSpec::new("Pair", group_anchor)
            .with_child(
                XlsxShapeSpec::shape("Left", child_anchor, XlsxShapePreset::Rectangle, "L").into(),
            )
            .with_child(
                XlsxShapeSpec::shape("Right", child_anchor, XlsxShapePreset::Ellipse, "R").into(),
            );
        group.transform = Some(XlsxGroupTransform {
            offset: Some(XlsxEmuOffset {
                x: XlsxEmu(0),
                y: XlsxEmu(0),
            }),
            extent: Some(XlsxEmuExtent {
                width: XlsxEmu(4_000_000),
                height: XlsxEmu(2_000_000),
            }),
            child_offset: Some(XlsxEmuOffset {
                x: XlsxEmu(0),
                y: XlsxEmu(0),
            }),
            child_extent: Some(XlsxEmuExtent {
                width: XlsxEmu(4_000_000),
                height: XlsxEmu(2_000_000),
            }),
        });

        let connection = XlsxConnectionShapeSpec::new(
            "Bridge",
            XlsxShapeAnchor::Absolute {
                position: XlsxEmuOffset {
                    x: XlsxEmu(500_000),
                    y: XlsxEmu(500_000),
                },
                extent: XlsxEmuExtent {
                    width: XlsxEmu(1_000_000),
                    height: XlsxEmu(1_000_000),
                },
            },
            XlsxShapePreset::StraightConnector1,
            XlsxConnectionEndSpec {
                shape_name: "Left".to_string(),
                site: 1,
            },
            XlsxConnectionEndSpec {
                shape_name: "Right".to_string(),
                site: 2,
            },
        );

        let mut sheet = MutableXlsbWorksheet::new("Shapes");
        sheet.add_shape(standalone).unwrap();
        sheet.add_group(group).unwrap();
        sheet.add_connection(connection).unwrap();
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(sheet);
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();

        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        let drawing = reader.sheet_drawing(0).expect("shape drawing missing");
        assert!(drawing.images.is_empty());
        assert!(drawing.charts.is_empty());
        assert_eq!(drawing.drawing.anchors.len(), 3);
        assert_eq!(drawing.shapes.len(), 3);

        let XlsxDrawingObject::Shape(shape) = &drawing.shapes[0].object else {
            panic!("expected standalone shape");
        };
        assert_eq!(shape.non_visual.id, Some(1));
        assert_eq!(shape.non_visual.name.as_deref(), Some("Standalone"));
        assert_eq!(
            shape.non_visual.description.as_deref(),
            Some("Typed XLSB text box")
        );
        assert_eq!(shape.text_body.as_ref().unwrap().text(), "A\nB");
        assert_eq!(
            shape.text_body.as_ref().unwrap().paragraphs[0].runs[0].bold,
            Some(true)
        );

        let XlsxDrawingObject::Group(group) = &drawing.shapes[1].object else {
            panic!("expected shape group");
        };
        assert_eq!(group.non_visual.id, Some(2));
        assert_eq!(group.children.len(), 2);
        let XlsxDrawingObject::Shape(left) = &group.children[0] else {
            panic!("expected left group child");
        };
        let XlsxDrawingObject::Shape(right) = &group.children[1] else {
            panic!("expected right group child");
        };
        assert_eq!(left.non_visual.id, Some(3));
        assert_eq!(right.non_visual.id, Some(4));

        let XlsxDrawingObject::ConnectionShape(connection) = &drawing.shapes[2].object else {
            panic!("expected connection shape");
        };
        assert_eq!(connection.non_visual.id, Some(5));
        assert_eq!(connection.start.unwrap().shape_id, 3);
        assert_eq!(connection.end.unwrap().shape_id, 4);

        let drawing_part = reader
            .opc_package()
            .iter_parts()
            .find(|part| part.content_type() == ct::OFC_DRAWING)
            .unwrap();
        assert!(drawing_part.rels().is_empty());
    }

    #[test]
    fn worksheet_shape_crud_and_save_validation_are_lossless_or_refuse() {
        use crate::xlsx::writer::{
            XlsxConnectionEndSpec, XlsxConnectionShapeSpec, XlsxGroupSpec, XlsxShapeSpec,
        };
        use crate::xlsx::{XlsxCellMarker, XlsxEditAs, XlsxEmu, XlsxShapeAnchor, XlsxShapePreset};

        fn anchor(from: (u32, u32), to: (u32, u32)) -> XlsxShapeAnchor {
            let marker = |(column, row)| XlsxCellMarker {
                column,
                row,
                column_offset: XlsxEmu(0),
                row_offset: XlsxEmu(0),
            };
            XlsxShapeAnchor::TwoCell {
                from: marker(from),
                to: marker(to),
                edit_as: XlsxEditAs::TwoCell,
            }
        }

        let valid = XlsxShapeSpec::shape(
            "Target",
            anchor((0, 0), (2, 2)),
            XlsxShapePreset::Rectangle,
            "target",
        );
        let invalid = XlsxShapeSpec::shape(
            "Descending",
            anchor((2, 2), (1, 1)),
            XlsxShapePreset::Rectangle,
            "",
        );
        let mut sheet = MutableXlsbWorksheet::new("Shapes");
        assert!(sheet.add_shape(invalid).is_err());
        assert!(sheet.shapes().is_empty());
        let mut invalid_xml = valid.clone();
        invalid_xml.name = "invalid\u{0}name".to_string();
        assert!(sheet.add_shape(invalid_xml).is_err());
        let mut invalid_text_properties = valid.clone();
        invalid_text_properties.paragraphs[0].runs[0].font_size_hundredths = Some(0);
        invalid_text_properties.body_properties.column_count = 17;
        assert!(sheet.add_shape(invalid_text_properties).is_err());
        assert!(sheet.shapes().is_empty());
        sheet.add_shape(valid.clone()).unwrap();
        sheet
            .add_group(
                XlsxGroupSpec::new("Group", anchor((3, 3), (6, 6))).with_child(
                    XlsxShapeSpec::shape(
                        "Nested",
                        anchor((3, 3), (4, 4)),
                        XlsxShapePreset::Ellipse,
                        "",
                    )
                    .into(),
                ),
            )
            .unwrap();
        sheet
            .add_connection(XlsxConnectionShapeSpec::new(
                "Dangling",
                anchor((1, 1), (2, 2)),
                XlsxShapePreset::StraightConnector1,
                XlsxConnectionEndSpec {
                    shape_name: "Missing".to_string(),
                    site: 0,
                },
                XlsxConnectionEndSpec {
                    shape_name: "Target".to_string(),
                    site: 0,
                },
            ))
            .unwrap();

        let mut workbook = XlsbWorkbookWriter::new();
        workbook.add_worksheet(sheet);
        assert!(workbook.save(&mut Cursor::new(Vec::new())).is_err());
        let sheet = workbook.get_worksheet_mut(0).unwrap();
        assert_eq!(sheet.shapes().len(), 1);
        assert_eq!(sheet.groups().len(), 1);
        assert_eq!(sheet.connections().len(), 1);
        assert!(sheet.remove_shape(4).is_err());
        assert!(sheet.remove_group(4).is_err());
        assert!(sheet.remove_connection(4).is_err());
        sheet.remove_connection(0).unwrap();
        assert!(workbook.save(&mut Cursor::new(Vec::new())).is_ok());
        let sheet = workbook.get_worksheet_mut(0).unwrap();
        sheet.clear_drawing_shapes();
        assert!(sheet.shapes().is_empty());
        assert!(sheet.groups().is_empty());
        assert!(sheet.connections().is_empty());
        sheet.add_shape(valid).unwrap();
        assert_eq!(sheet.remove_shape(0).unwrap().name, "Target");
    }
}
