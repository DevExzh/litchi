use super::model::{CellValue, Writer};
use super::stream;
use crate::encryption::encrypt_workbook_for_write;
use crate::error::Result;
use litchi_cfb::writer::OleWriter;

impl Writer {
    /// Save the XLS file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), Error>` - Success or error
    ///
    /// # Implementation Status
    ///
    /// ✅ Basic structure generation (BOF, EOF, workbook globals)
    /// ✅ Cell record generation (Number, `LabelSST`, `BoolErr`)
    /// ✅ Shared string table (SST)
    /// ✅ Formula tokenization for the supported BIFF8 writer subset
    /// ❌ Cell formatting (XF records)
    /// ❌ Column widths / row heights
    /// ❌ Merged cells
    /// ❌ Named ranges
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(&mut ole_writer, streams)?;

        // Save to file
        ole_writer.save(path)?;

        Ok(())
    }

    /// Write to a writer (useful for testing and in-memory generation)
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer
    ///
    /// # Returns
    ///
    /// * `Result<(), Error>` - Success or error
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_to<W: std::io::Write + std::io::Seek>(&mut self, writer: &mut W) -> Result<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(&mut ole_writer, streams)?;

        // Write to the provided writer
        ole_writer.write_to(writer)?;

        Ok(())
    }

    fn populate_compound_document(
        &self,
        ole_writer: &mut OleWriter,
        streams: stream::WorkbookStreams,
    ) -> Result<()> {
        let stream::WorkbookStreams {
            workbook,
            toolbar,
            pivot_caches,
        } = streams;

        ole_writer.create_stream_owned(&["Workbook"], workbook)?;
        if let Some(metadata) = &self.vba_metadata {
            ole_writer.create_storage(&["_VBA_PROJECT_CUR"])?;
            metadata
                .project
                .write_into(ole_writer, &["_VBA_PROJECT_CUR"])?;
        }

        // Pivot cache storage: _SX_DB_CUR/XXXX. Stream names use four-digit
        // uppercase hexadecimal identifiers, matching the legacy convention.
        if !pivot_caches.is_empty() {
            ole_writer.create_storage(&["_SX_DB_CUR"])?;
            for (id, data) in pivot_caches {
                let name = format!("{id:04X}");
                ole_writer.create_stream_owned(&["_SX_DB_CUR", &name], data)?;
            }
        }
        if let Some(toolbar) = toolbar {
            ole_writer.create_stream_owned(&["XCB"], toolbar)?;
        }
        if let Some(map_info) = &self.xml_map {
            let xml = crate::xml_map::write(map_info)?;
            ole_writer.create_stream_owned(&[crate::xml_map::STREAM_NAME], xml)?;
        }
        Ok(())
    }

    /// Build the shared string table from all string cells
    pub(super) fn build_shared_strings(&mut self) {
        self.shared_strings.clear();
        self.string_map.clear();
        self.sst_total = 0;

        // Collect all unique strings from all worksheets
        for worksheet in &self.worksheets {
            for cell in worksheet.cells.values() {
                if let CellValue::String(ref s) = cell.value {
                    // Count total occurrences
                    self.sst_total = self.sst_total.saturating_add(1);
                    // Insert unique strings
                    if !self.string_map.contains_key(s) {
                        let index = crate::utils::truncate_usize_to_u32(self.shared_strings.len());
                        self.string_map.insert(s.clone(), index);
                        self.shared_strings.push(s.clone());
                    }
                }
            }
        }
    }

    /// Generate the complete Workbook stream (plus pivot cache streams) with
    /// all BIFF records.
    fn generate_workbook_streams(&self) -> Result<stream::WorkbookStreams> {
        if let Some(map_info) = self.xml_map.as_ref() {
            crate::xml_map::validate_info(map_info)?;
        }
        crate::xml_map::validate_list_objects(
            self.xml_map.as_ref(),
            self.worksheets
                .iter()
                .flat_map(|worksheet| worksheet.list_objects.iter()),
        )?;
        let mut streams = stream::generate_workbook_stream(
            self.use_1904_dates,
            self.calculation_settings,
            self.vba_metadata.as_ref(),
            self.environment_options,
            self.workbook_window_options,
            &self.function_group_options,
            &self.external_workbooks,
            &self.external_names,
            &self.add_in_functions,
            &self.dde_or_ole_links,
            &self.fmt,
            self.custom_table_styles.as_ref(),
            &self.defined_names,
            &self.defined_name_records,
            &self.shared_strings,
            self.sst_total,
            self.workbook_protection,
            self.file_sharing.as_ref(),
            self.book_ext.as_ref(),
            self.theme.as_ref(),
            self.mdx_metadata.as_ref(),
            &self.real_time_data,
            &self.web_publications,
            &self.xf_extensions,
            &self.style_extensions,
            &self.worksheets,
            &self.string_map,
        )?;
        streams.toolbar = self
            .toolbar
            .as_ref()
            .map(crate::toolbar::to_bytes)
            .transpose()?;
        if let Some(encryption) = &self.encryption {
            streams.workbook = encrypt_workbook_for_write(streams.workbook, encryption)?;
        }
        Ok(streams)
    }

    /// Get the number of worksheets in this workbook
    #[must_use]
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get worksheet name by index
    #[must_use]
    pub fn get_worksheet_name(&self, index: usize) -> Option<&str> {
        self.worksheets.get(index).map(|w| w.name.as_str())
    }

    // Implementation status notes:
    // ✅ Building shared string table (SST) with deduplication - IMPLEMENTED
    // ✅ Generating BIFF8 records for supported cell types - Number, LabelSST, BoolErr, Formula
    // ❌ Worksheet management (rename, delete, reorder) - Future enhancement
    // ❌ Cell formatting (fonts, colors, borders, number formats) - Future enhancement
    // ❌ Column widths and row heights - Future enhancement
    // ❌ Merged cells - Future enhancement
    // ✅ Named ranges (simple A1-style, workbook and sheet scoped) - IMPLEMENTED
    // ✅ Formula parsing and tokenization for the supported writer subset
}
