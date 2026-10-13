use super::model::Builder;
use super::stream;
use crate::encryption::{EncryptionEntropy, encrypt_workbook_for_write};
use crate::error::{Error, Result};
use litchi_cfb::writer::OleWriter;

impl Builder {
    /// Build a validated immutable XLS package from the staged workbook.
    ///
    /// An empty builder receives one blank `Sheet1`. Subsequent builds retain
    /// the staged content and can be published independently. Encrypted builds
    /// retain their encrypted image; publish that image explicitly with
    /// [`crate::Package::to_encrypted_bytes`]. Ordinary package publication
    /// refuses encrypted owners. Office Binary RC4 and CryptoAPI RC4 builds
    /// also require a one-use [`crate::EncryptionEntropy`] supplied through
    /// [`Builder::set_encryption_entropy`](crate::Builder::set_encryption_entropy);
    /// the entropy is consumed even when this method later fails.
    ///
    /// # Errors
    /// Returns typed authoring, container or strict package-admission failures.
    pub fn build(&mut self) -> crate::package::Result<crate::Package> {
        // Take caller-owned RC4 entropy before any default-sheet staging.  A
        // missing token is therefore a typed preflight error and cannot leave
        // a visible mutation behind.  Taking it also ensures that all RC4
        // attempts consume the token, including attempts that fail later in
        // generation or strict package admission.
        let encryption_entropy = match self
            .encryption
            .as_ref()
            .map(|encryption| encryption.profile)
        {
            Some(crate::EncryptionProfile::OfficeBinaryRc4)
            | Some(crate::EncryptionProfile::CryptoApiRc4 { .. }) => Some(
                self.encryption_entropy
                    .take()
                    .ok_or(Error::EncryptionEntropyRequired)?,
            ),
            _ => None,
        };
        let properties = self.property_streams()?;
        if self.worksheets.is_empty() {
            self.add_worksheet("Sheet1")?;
        }
        let streams = self.generate_workbook_streams(encryption_entropy.as_ref())?;
        let mut ole_writer = OleWriter::new();
        self.populate_compound_document(&mut ole_writer, streams)?;
        if let Some(properties) = properties {
            properties.write_into(&mut ole_writer)?;
        }
        let mut output = std::io::Cursor::new(Vec::new());
        ole_writer.write_to(&mut output)?;
        let password = self
            .encryption
            .as_ref()
            .map(|encryption| encryption.password.as_str());
        crate::Package::from_owned_bytes(
            output.into_inner(),
            crate::OpenOptions::new().password(password),
        )
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

    /// Stage the shared string table from all string cells, borrowing them.
    ///
    /// # Errors
    ///
    /// Refuses a string that an SST entry cannot hold.
    pub(super) fn shared_string_table(&self) -> Result<stream::SharedStringTable<'_>> {
        stream::SharedStringTable::build(&self.worksheets)
    }

    /// Generate the complete Workbook stream (plus pivot cache streams) with
    /// all BIFF records.
    fn generate_workbook_streams(
        &self,
        encryption_entropy: Option<&EncryptionEntropy>,
    ) -> Result<stream::WorkbookStreams> {
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
            &self.shared_string_table()?,
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
        )?;
        streams.toolbar = self
            .toolbar
            .as_ref()
            .map(crate::toolbar::to_bytes)
            .transpose()?;
        if let Some(encryption) = &self.encryption {
            streams.workbook =
                encrypt_workbook_for_write(streams.workbook, encryption, encryption_entropy)?;
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
