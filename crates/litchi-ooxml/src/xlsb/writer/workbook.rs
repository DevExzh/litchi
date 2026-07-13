//! XLSB workbook writer implementation
//!
//! This module provides functionality to create complete XLSB files with multiple worksheets,
//! shared strings, styles, and advanced features.
use crate::xlsb::calculation::CalculationProperties;
use crate::xlsb::error::XlsbResult;
use crate::xlsb::formula::{
    CellParsedFormula, FormulaCompilationContext, FormulaDefinedName, excel_name_eq,
};
use crate::xlsb::named_ranges::{NamedRange, validate_defined_name};
use crate::xlsb::records::record_types;
use crate::xlsb::writer::{
    MutableSharedStringsWriter, MutableXlsbWorksheet, RecordWriter, StylesWriter,
};
use litchi_core::xml::escape_xml;
use litchi_opc::constants::relationship_type as rel;
use litchi_opc::part::Part;
use litchi_opc::{BlobPart, OpcPackage, PackURI};
use std::io::{Seek, Write};

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
    named_ranges: Vec<NamedRange>,
    shared_strings: MutableSharedStringsWriter,
    styles: StylesWriter,
    calculation_properties: CalculationProperties,
    is_1904: bool,
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
            named_ranges: Vec::new(),
            shared_strings: MutableSharedStringsWriter::new(),
            styles: StylesWriter::new(),
            calculation_properties: CalculationProperties::default(),
            is_1904: false,
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

    /// Workbook formula calculation policy written to `BrtCalcProp`.
    pub fn calculation_properties(&self) -> &CalculationProperties {
        &self.calculation_properties
    }

    /// Mutably configure workbook formula calculation policy.
    pub fn calculation_properties_mut(&mut self) -> &mut CalculationProperties {
        &mut self.calculation_properties
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
    }

    /// Add a named range (defined name) to the workbook.
    pub fn add_named_range(&mut self, named_range: NamedRange) {
        self.named_ranges.push(named_range);
    }

    /// Get a mutable reference to a worksheet by index
    pub fn get_worksheet_mut(&mut self, index: usize) -> Option<&mut MutableXlsbWorksheet> {
        self.worksheets.get_mut(index)
    }

    /// Get the number of worksheets
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
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

        // Save package to output
        package.to_stream(writer)?;

        Ok(())
    }

    fn validate_formula_metadata(&self) -> XlsbResult<()> {
        if self.worksheets.len() > usize::from(u16::MAX) - 2 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "{} worksheets exceed the XLSB extern-sheet index limit",
                self.worksheets.len()
            )));
        }
        for (index, worksheet) in self.worksheets.iter().enumerate() {
            let name = worksheet.name();
            let name_len = name.encode_utf16().count();
            if name_len == 0
                || name_len > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "worksheet name {name:?} does not follow BrtBundleSh grammar"
                )));
            }
            if self.worksheets[..index]
                .iter()
                .any(|existing| excel_name_eq(existing.name(), name))
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "duplicate worksheet name {name:?}"
                )));
            }
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
                    .is_none_or(|sheet_id| sheet_id >= self.worksheets.len())
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

    /// Write workbook-level defined names (BrtName records).
    fn write_named_ranges<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        for named_range in &self.named_ranges {
            if named_range.function {
                return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                    "macro defined name {} cannot be emitted",
                    named_range.name
                )));
            }
            validate_defined_name(&named_range.name)?;
            if let Some(sheet_id) = named_range.sheet_id {
                if usize::try_from(sheet_id)
                    .ok()
                    .is_none_or(|index| index >= self.worksheets.len())
                {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "defined name {} has invalid sheet scope {sheet_id}",
                        named_range.name
                    )));
                }
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
            let mut temp_writer = RecordWriter::new(&mut data);

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

            writer.write_record(record_types::NAME, &data)?;
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
        let sheet_count = self.worksheets.len();

        // Build sheet names list
        let mut sheet_names = String::new();
        for sheet in &self.worksheets {
            sheet_names.push_str(&format!(
                "<vt:lpstr>{}</vt:lpstr>",
                escape_xml(sheet.name())
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
        let mut workbook_data = Vec::new();
        let mut writer = RecordWriter::new(&mut workbook_data);

        // Write workbook structure
        self.write_workbook(&mut writer, formula_sheet_ranges)?;

        // Create workbook part
        let workbook_uri = PackURI::new("/xl/workbook.bin")?;
        let mut workbook_part = BlobPart::new(
            workbook_uri.clone(),
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
            workbook_data,
        );

        // Add relationships from workbook to worksheets and styles
        {
            let rels = workbook_part.rels_mut();
            for i in 0..self.worksheets.len() {
                rels.get_or_add(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    &format!("worksheets/sheet{}.bin", i + 1),
                );
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
        }

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
    /// BrtBeginExternals / BrtSupSelf / BrtExternSheet / BrtEndExternals
    /// [BrtCalcProp]
    /// BrtEndBook (0x0084)
    /// ```
    ///
    /// The book views and calculation properties are currently written with a
    /// single default view and sensible defaults for calculation settings.
    fn write_workbook<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        formula_sheet_ranges: &[(u32, u32)],
    ) -> XlsbResult<()> {
        // BrtBeginBook
        writer.write_record(record_types::BEGIN_BOOK, &[])?;

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

        // EXTERNALS block with self-references, mirroring SheetJS and
        // [MS-XLSB] examples. This creates a minimal but fully valid
        // extern sheet table for the workbook.
        self.write_externals(writer, formula_sheet_ranges)?;

        // Defined names (named ranges), if any.
        self.write_named_ranges(writer)?;

        // Basic calculation properties describing recalc behavior and
        // numerical tolerance. This is tiny and follows the spec example
        // values, so we emit it unconditionally.
        self.write_calc_properties(writer)?;

        // BrtEndBook
        writer.write_record(record_types::END_BOOK, &[])?;

        Ok(())
    }

    /// Write file version record (BrtFileVersion)
    /// This is REQUIRED for Excel to open the file
    fn write_file_version<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // Build structure per spec example (48 bytes total):
        // guidCodeName (16 zero bytes), stAppName ("xl"), stLastEdited ("4"),
        // stLowestEdited ("4"), stRupBuild ("4505")
        let mut data = Vec::with_capacity(48);
        let mut w = RecordWriter::new(&mut data);

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

        writer.write_record(record_types::FILE_VERSION, &data)?;
        Ok(())
    }

    /// Write workbook properties (BrtWbProp)
    fn write_workbook_properties<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

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

        writer.write_record(record_types::WORKBOOK_PROP, &data)?;
        Ok(())
    }

    /// Write book views (REQUIRED by Excel)
    fn write_book_views<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        writer.write_record(record_types::BEGIN_BOOK_VIEWS, &[])?;

        // Write one default book view
        let mut view_data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut view_data);

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

        writer.write_record(record_types::BOOK_VIEW, &view_data)?;

        writer.write_record(record_types::END_BOOK_VIEWS, &[])?;
        Ok(())
    }

    /// Write bundle sheets (worksheet metadata)
    fn write_bundle_sheets<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        writer.write_record(record_types::BEGIN_BUNDLE_SHS, &[])?;

        for (i, worksheet) in self.worksheets.iter().enumerate() {
            let mut sheet_data = Vec::new();
            let mut temp_writer = RecordWriter::new(&mut sheet_data);

            // hsState (u32): 0 = visible
            temp_writer.write_u32(0)?;
            // itabID (u32): unique sheet id (1-based)
            temp_writer.write_u32((i + 1) as u32)?;
            // RelID (XLWideString): rIdN
            temp_writer.write_wide_string(&format!("rId{}", i + 1))?;
            // strName (XLWideString): sheet name
            temp_writer.write_wide_string(worksheet.name())?;

            writer.write_record(record_types::BUNDLE_SH, &sheet_data)?;
        }

        writer.write_record(record_types::END_BUNDLE_SHS, &[])?;
        Ok(())
    }

    /// Write calculation properties (CALC_PROP, 0x009D)
    ///
    /// Spec example fields and order
    fn write_calc_properties<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        self.calculation_properties.validate()?;
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        temp_writer.write_u32(self.calculation_properties.recalculation_id)?;
        temp_writer.write_u32(self.calculation_properties.mode as u32)?;
        temp_writer.write_u32(self.calculation_properties.iteration_count)?;
        temp_writer.write_f64(self.calculation_properties.iteration_delta)?;
        temp_writer.write_u32(self.calculation_properties.user_thread_count as u32)?;
        temp_writer.write_u16(self.calculation_properties.flags())?;

        writer.write_record(record_types::CALC_PROP, &data)?;
        Ok(())
    }

    /// Write externals section (self-references)
    ///
    /// Based on SheetJS implementation: always writes BrtSupSelf with BrtExternSheet
    /// This creates self-references for the workbook and all sheets.
    fn write_externals<W: Write>(
        &self,
        writer: &mut RecordWriter<W>,
        formula_sheet_ranges: &[(u32, u32)],
    ) -> XlsbResult<()> {
        // BrtBeginExternals - no data
        writer.write_record(record_types::BEGIN_EXTERNALS, &[])?;

        // BrtSupSelf - no data
        writer.write_record(record_types::SUP_SELF, &[])?;

        // BrtExternSheet - self-references data
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        let sheet_count = self.worksheets.len();

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

        writer.write_record(record_types::EXTERN_SHEET, &data)?;

        // BrtEndExternals - no data
        writer.write_record(record_types::END_EXTERNALS, &[])?;

        Ok(())
    }

    /// Add worksheet parts to the package
    fn add_worksheet_parts(&mut self, package: &mut OpcPackage) -> XlsbResult<Vec<(u32, u32)>> {
        let worksheet_names = self
            .worksheets
            .iter()
            .map(|worksheet| worksheet.name().to_string())
            .collect::<Vec<_>>();
        let defined_names = self
            .named_ranges
            .iter()
            .map(|named_range| FormulaDefinedName {
                name: named_range.name.clone(),
                sheet_id: named_range.sheet_id,
            })
            .collect::<Vec<_>>();
        let formula_sheet_ranges = std::cell::RefCell::new(Vec::new());
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
                    &mut RecordWriter::new(&mut comments_data),
                    worksheet.comments(),
                )?;
                Some(BlobPart::new(
                    PackURI::new(format!("/xl/{comments_name}"))?,
                    "application/vnd.ms-excel.comments".to_string(),
                    comments_data,
                ))
            };

            // Now serialize the worksheet with fully-populated relationship IDs
            // in the hyperlink records.
            let mut sheet_data = Vec::new();
            let current_sheet = u32::try_from(i).map_err(|_| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "worksheet index overflow".to_string(),
                )
            })?;
            let formula_context = FormulaCompilationContext {
                worksheet_names: &worksheet_names,
                defined_names: &defined_names,
                sheet_ranges: &formula_sheet_ranges,
                current_sheet,
            };
            let compiled_formulas = worksheet.compile_contextual_formulas(&formula_context)?;
            let write_result = {
                let mut writer = RecordWriter::new(&mut sheet_data);
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
        }

        Ok(formula_sheet_ranges.into_inner())
    }

    /// Add shared strings part to the package
    fn add_shared_strings_part(&self, package: &mut OpcPackage) -> XlsbResult<()> {
        let mut sst_data = Vec::new();
        let mut writer = RecordWriter::new(&mut sst_data);

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
        let mut writer = RecordWriter::new(&mut styles_data);

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
    use crate::xlsb::{CalculationMode, SharedStringRun, SheetProtection};
    use litchi_core::sheet::{CellValue, WorkbookTrait};
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
    fn test_set_date_system() {
        let mut workbook = XlsbWorkbookWriter::new();
        workbook.set_date_system(true);
        assert!(workbook.is_1904);
    }

    #[test]
    fn calculation_properties_survive_package_roundtrip() {
        let mut workbook = XlsbWorkbookWriter::new();
        let properties = workbook.calculation_properties_mut();
        properties.mode = CalculationMode::Manual;
        properties.iterative_calculation = true;
        properties.iteration_count = 25;
        properties.iteration_delta = 0.000_01;
        properties.user_set_thread_count = true;
        properties.user_thread_count = 4;
        properties.full_calculation_on_load = true;
        workbook.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));

        let expected = workbook.calculation_properties().clone();
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).unwrap();
        let reader = crate::xlsb::XlsbWorkbook::new(Cursor::new(output.into_inner())).unwrap();
        assert_eq!(reader.calculation_properties(), &expected);
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
}
