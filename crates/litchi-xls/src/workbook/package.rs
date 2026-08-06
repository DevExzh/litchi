//! CFB package and transaction orchestration for XLS workbooks.
//!
//! This layer owns opening the MS-XLS OLE compound file, selecting/decrypting
//! the Workbook stream, and accessing package-level property, signature,
//! VBA, custom-XML, and revision-log streams. BIFF grammar stays in the
//! sibling codec module.

use super::codec::{WorkbookGlobalsSink, pivot_cache_stream_paths};
use super::model::{OpenOptions, Workbook};
use crate::defined_names::DefinedNameSlot;
use crate::encryption::prepare_workbook_stream;
use crate::error::{Error, Result};
use crate::leniency::{ToleranceLog, ToleranceReport};
use crate::number_format::Formatting;
use crate::pivot_table;
use crate::records::{BiffVersion, Encoding};
use crate::sheet_metadata::SheetMetadata;
use litchi_biff::Records;
use litchi_cfb::OleFile;
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::collections::HashSet;
use std::io::{Read, Seek};
use std::sync::Arc;

impl<R: Read + Seek> Workbook<R> {
    fn empty(ole_file: OleFile<R>) -> Self {
        Self {
            ole_file,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            sheets: Vec::new(),
            shared_strings: None,
            shared_string_properties: None,
            shared_string_reference_count: 0,
            palette: crate::palette::Palette::default(),
            fonts: Vec::new(),
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: crate::formula::FormulaContext::default(),
            external_links: crate::Links::default(),
            pivot_caches: Vec::new(),
            pivot_cache_stream_ids: Vec::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            formatting: Arc::new(Formatting::default()),
            protection: crate::protection::WorkbookProtection::default(),
            calculation: crate::calculation::WorkbookCalculation::default(),
            vba_metadata: crate::vba::VbaMetadata::default(),
            environment: crate::environment::WorkbookEnvironment::default(),
            book_ext: None,
            style_extensions: Vec::new(),
            theme: None,
            write_access: Ok(None),
            table_styles: None,
            shared_string_index: Ok(None),
            workbook_view: crate::workbook_view::WorkbookView::default(),
            custom_views: Vec::new(),
            real_time_data: Vec::new(),
            mdx_metadata: crate::mdx_metadata::MdxMetadata::default(),
            web_publications: Vec::new(),
            function_groups: None,
            tolerance: ToleranceReport::default(),
        }
    }

    pub fn new(reader: R) -> Result<Self> {
        Self::new_with_options(reader, OpenOptions::default())
    }

    /// Open an XLS workbook with an explicit password contract.
    pub fn new_with_options(reader: R, options: OpenOptions<'_>) -> Result<Self> {
        let mut workbook = Self::empty(OleFile::open(reader)?);

        workbook.parse_workbook(&options)?;
        Ok(workbook)
    }

    /// Create an XLS workbook from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - An already-parsed OLE file
    pub fn from_ole_file(ole_file: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_options(ole_file, OpenOptions::default())
    }

    /// Create a workbook from a parsed OLE file with explicit open options.
    pub fn from_ole_file_with_options(
        ole_file: OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let mut workbook = Self::empty(ole_file);

        workbook.parse_workbook(&options)?;
        Ok(workbook)
    }

    /// Parse the workbook stream
    fn parse_workbook(&mut self, options: &OpenOptions<'_>) -> Result<()> {
        // Find and read the Workbook stream
        let workbook_data = self
            .ole_file
            .open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;
        let workbook_data = prepare_workbook_stream(workbook_data, options.password)?;
        let mut tolerance = ToleranceLog::new(options.leniency);

        let mut records = Records::new(&workbook_data);
        let mut encoding = Encoding::from_codepage(1252)?; // Default codepage
        let mut bound_sheets = Vec::new();
        let mut strings = Vec::new();
        let mut string_properties = Vec::new();

        // Parse workbook globals
        let mut defined_name_slots = Vec::new();
        self.parse_workbook_globals(
            &mut records,
            &mut encoding,
            WorkbookGlobalsSink {
                bound_sheets: &mut bound_sheets,
                strings: &mut strings,
                string_properties: &mut string_properties,
                defined_name_slots: &mut defined_name_slots,
                tolerance: &mut tolerance,
            },
        )?;
        self.tolerance = tolerance.into_report();

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.shared_string_properties = Some(Arc::new(string_properties));
        let all_sheet_names = bound_sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let mut unique_sheet_names = HashSet::with_capacity(all_sheet_names.len());
        for name in &all_sheet_names {
            if !unique_sheet_names.insert(name.to_lowercase()) {
                return Err(Error::InvalidRecord {
                    record_type: 0x0085,
                    message: format!("duplicate case-insensitive BoundSheet8 name: {name:?}"),
                });
            }
        }
        self.sheets = bound_sheets
            .iter()
            .enumerate()
            .map(|(index, sheet)| SheetMetadata::from_bound_sheet(index, sheet))
            .collect();
        self.worksheet_names.clear();
        self.formula_context.set_sheet_names(all_sheet_names);
        self.formula_context.set_scoped_defined_names(
            defined_name_slots
                .iter()
                .map(DefinedNameSlot::formula_symbol)
                .collect(),
        );
        self.formula_context
            .set_external_links(&self.external_links);
        self.defined_name_records = defined_name_slots
            .into_iter()
            .map(|slot| slot.into_public(bound_sheets.len(), &self.formula_context))
            .collect::<Result<Vec<_>>>()?;
        self.defined_names = self
            .defined_name_records
            .iter()
            .filter(|name| !name.is_macro())
            .cloned()
            .collect();

        // Parse worksheets from positions in the workbook stream
        for (sheet_index, bound_sheet) in bound_sheets.iter().enumerate() {
            if bound_sheet.sheet_type != crate::records::SheetType::WorkSheet {
                continue;
            }
            match self.parse_worksheet_from_position(&workbook_data, bound_sheet, &encoding) {
                Ok(worksheet) => {
                    let worksheet_index = self.worksheets.len();
                    self.worksheet_names.push(bound_sheet.name.clone());
                    self.worksheets.push(worksheet);
                    self.sheets[sheet_index].set_parsed_worksheet_index(worksheet_index);
                },
                Err(error @ Error::InvalidRecord { record_type, .. })
                    if pivot_table::is_worksheet_view_record(record_type) =>
                {
                    return Err(error);
                },
                Err(_) => {},
            }
        }

        let cache_paths = pivot_cache_stream_paths(self.ole_file.list_streams());
        let mut pivot_caches = Vec::with_capacity(cache_paths.len());
        for (expected_stream_id, path) in cache_paths {
            if expected_stream_id == 0 {
                return Err(Error::InvalidRecord {
                    record_type: 0x00C6,
                    message: "PivotCache storage stream ID must be nonzero".to_string(),
                });
            }
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = self.ole_file.open_stream(&refs)?;
            let cache = pivot_table::parse_pivot_cache_stream(&data)?;
            if cache.stream_id() != expected_stream_id {
                return Err(Error::InvalidRecord {
                    record_type: 0x00C6,
                    message: format!(
                        "PivotCache storage stream {:04X} contains stream ID {:04X}",
                        expected_stream_id,
                        cache.stream_id()
                    ),
                });
            }
            pivot_caches.push(cache);
        }
        self.pivot_caches = pivot_caches;
        pivot_table::validate_pivot_cache_links(
            &self.worksheets,
            &self.pivot_caches,
            &self.pivot_cache_stream_ids,
        )?;

        let visible_tabs = bound_sheets
            .iter()
            .map(|sheet| matches!(sheet.visible, crate::records::SheetVisible::Visible))
            .collect::<Vec<_>>();
        let selected_worksheet_tabs = self
            .sheets
            .iter()
            .map(|sheet| {
                sheet
                    .parsed_worksheet_index()
                    .and_then(|index| self.worksheets.get(index))
                    .and_then(|worksheet| worksheet.worksheet_view())
                    .map(|view| view.is_selected())
            })
            .collect::<Vec<_>>();
        self.workbook_view
            .validate_sheet_state(&visible_tabs, &selected_worksheet_tabs)?;

        Ok(())
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    pub fn custom_xml_data_store(
        &mut self,
    ) -> litchi_ole_common::custom_xml::Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole_file)
    }

    pub fn summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify workbook XML signatures with the safe strict policy, without
    /// evaluating certificate trust or executing any macro content.
    pub fn signatures(&mut self) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        self.signatures_with(&litchi_sign::Policy::strict())
    }

    /// Verify workbook XML signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole_file, litchi_sign::cfb::Format::Xls, policy)
    }

    pub fn document_summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Read the optional Office Toolbars (`XCB`) stream as inert typed metadata.
    ///
    /// The returned value owns decoded strings and visual bytes because the
    /// compound-file stream buffer is temporary. No command, macro, UI, or
    /// ActiveX behavior is activated while reading this stream.
    pub fn toolbar(&mut self) -> Result<Option<crate::Wrapper<'static>>> {
        match self.ole_file.open_stream(&["XCB"]) {
            Ok(data) => crate::toolbar::parse(&data)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidData(format!("XCB toolbar: {error}"))),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self
            .document_summary_information()?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }

    pub fn vba_metadata(&self) -> crate::vba::VbaMetadata {
        let mut metadata = self.vba_metadata.clone();
        metadata.set_project_storage_present(self.vba_project_storage().is_some());
        metadata
    }

    /// Discover the MS-XLS `_VBA_PROJECT_CUR` storage without opening macro streams.
    ///
    /// This validates directory names defined by MS-XLS and MS-OVBA only. It
    /// never opens, decompresses, parses, or executes `PROJECT`, `dir`,
    /// `_VBA_PROJECT`, SRP, or module-stream bytes.
    pub fn vba_project_storage(&self) -> Option<crate::VbaProjectStorage> {
        crate::vba::discover_vba_project_storage(&self.ole_file.list_streams())
    }

    /// Parse the `_VBA_PROJECT_CUR` MS-OVBA project with safe default limits.
    ///
    /// The method returns `None` when no structurally complete VBA project is
    /// present. Source is only decompressed and decoded; it is never compiled,
    /// interpreted, or executed.
    pub fn vba(
        &mut self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        self.vba_with(&litchi_vba::Limits::default())
    }

    /// Parse the `_VBA_PROJECT_CUR` project with explicit resource limits.
    pub fn vba_with(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        let Some(storage) = self.vba_project_storage() else {
            return Ok(None);
        };
        if !storage.is_structurally_complete() {
            return Ok(None);
        }
        let path: Vec<&str> = storage
            .root_storage_path()
            .iter()
            .map(String::as_str)
            .collect();
        litchi_vba::project::Project::open(&mut self.ole_file, &path, limits).map(Some)
    }

    /// Whether the CFB container holds a shared-workbook `Revision Log` stream
    /// (MS-XLS 2.1.7.14).
    pub fn has_revision_log(&self) -> bool {
        crate::revision_log::find_revision_log_stream(&self.ole_file.list_streams()).is_some()
    }

    /// Parse the shared-workbook `Revision Log` stream, when present.
    ///
    /// The result is a typed, inert model of the RRD revision records.
    /// Parsing never applies, rejects, or replays any recorded revision.
    pub fn revision_log(&mut self) -> Result<Option<crate::revision_log::RevisionLog>> {
        let Some(name) =
            crate::revision_log::find_revision_log_stream(&self.ole_file.list_streams())
                .map(str::to_owned)
        else {
            return Ok(None);
        };
        let data = self.ole_file.open_stream(&[name.as_str()])?;
        crate::revision_log::parse_revision_log_stream(&data).map(Some)
    }
}
