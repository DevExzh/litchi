//! XLSB workbook writer implementation
//!
//! This module provides functionality to create complete XLSB files with multiple worksheets,
//! shared strings, styles, and advanced features.

use crate::ooxml::opc::constants::relationship_type as rel;
use crate::ooxml::opc::part::Part;
use crate::ooxml::opc::{BlobPart, OpcPackage, PackURI};
use crate::ooxml::xlsb::error::XlsbResult;
use crate::ooxml::xlsb::records::record_types;
use crate::ooxml::xlsb::writer::{
    MutableSharedStringsWriter, MutableXlsbWorksheet, RecordWriter, StylesWriter,
};
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
/// use litchi::ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
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
    shared_strings: MutableSharedStringsWriter,
    styles: StylesWriter,
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
            shared_strings: MutableSharedStringsWriter::new(),
            styles: StylesWriter::new(),
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

    /// Add a worksheet to the workbook
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ooxml::xlsb::writer::{XlsbWorkbookWriter, MutableXlsbWorksheet};
    ///
    /// let mut workbook = XlsbWorkbookWriter::new();
    /// let sheet = MutableXlsbWorksheet::new("Sheet1");
    /// workbook.add_worksheet(sheet);
    /// ```
    pub fn add_worksheet(&mut self, worksheet: MutableXlsbWorksheet) {
        self.worksheets.push(worksheet);
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
        let mut package = OpcPackage::new();

        // Add document properties (required by Excel)
        self.add_doc_props(&mut package)?;

        // Add theme (REQUIRED by Excel)
        self.add_theme(&mut package)?;

        // Add worksheets first so that shared_strings is fully populated before we
        // decide whether to create a sharedStrings part and relationship.
        self.add_worksheet_parts(&mut package)?;

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
        self.add_workbook_part(&mut package)?;

        // Save package to output
        package.to_stream(writer)?;

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
    fn add_workbook_part(&self, package: &mut OpcPackage) -> XlsbResult<()> {
        let mut workbook_data = Vec::new();
        let mut writer = RecordWriter::new(&mut workbook_data);

        // Write workbook structure
        self.write_workbook(&mut writer)?;

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
    fn write_workbook<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
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
        self.write_externals(writer)?;

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
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        // recalcID (DWORD)
        temp_writer.write_u32(0x0001_EB1D)?;
        // fAutoRecalc (LONG)
        temp_writer.write_u32(1)?;
        // cCalcCount (DWORD)
        temp_writer.write_u32(100)?;
        // xnumDelta (Xnum/f64): 0.001
        temp_writer.write_f64(0.001f64)?;
        // cUserThreadCount (LONG)
        temp_writer.write_u32(1)?;
        // Flags (WORD) with bits per spec: 0b0110_1010 = 0x006A
        temp_writer.write_u16(0x006A)?;

        writer.write_record(record_types::CALC_PROP, &data)?;
        Ok(())
    }

    /// Write externals section (self-references)
    ///
    /// Based on SheetJS implementation: always writes BrtSupSelf with BrtExternSheet
    /// This creates self-references for the workbook and all sheets.
    fn write_externals<W: Write>(&self, writer: &mut RecordWriter<W>) -> XlsbResult<()> {
        // BrtBeginExternals - no data
        writer.write_record(record_types::BEGIN_EXTERNALS, &[])?;

        // BrtSupSelf - no data
        writer.write_record(record_types::SUP_SELF, &[])?;

        // BrtExternSheet - self-references data
        let mut data = Vec::new();
        let mut temp_writer = RecordWriter::new(&mut data);

        let sheet_count = self.worksheets.len();

        // Total count: sheet_count + 2
        temp_writer.write_u32((sheet_count + 2) as u32)?;

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

        writer.write_record(record_types::EXTERN_SHEET, &data)?;

        // BrtEndExternals - no data
        writer.write_record(record_types::END_EXTERNALS, &[])?;

        Ok(())
    }

    /// Add worksheet parts to the package
    fn add_worksheet_parts(&mut self, package: &mut OpcPackage) -> XlsbResult<()> {
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

            // Now serialize the worksheet with fully-populated relationship IDs
            // in the hyperlink records.
            let mut sheet_data = Vec::new();
            {
                let mut writer = RecordWriter::new(&mut sheet_data);
                worksheet.write(&mut writer, &mut self.shared_strings)?;
            }
            sheet_part.set_blob(sheet_data);

            package.add_part(Box::new(sheet_part));
            package.add_part(Box::new(binary_index_part));
        }

        Ok(())
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

/// Escape XML special characters
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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
}
