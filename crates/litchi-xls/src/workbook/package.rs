//! CFB package and transaction orchestration for XLS workbooks.
//!
//! This layer owns opening the MS-XLS OLE compound file, selecting/decrypting
//! the Workbook stream, and accessing package-level property, signature,
//! VBA, custom-XML, and revision-log streams. BIFF grammar stays in the
//! sibling codec module.

#[path = "package/property_set.rs"]
pub mod property_set;

use super::codec::{WorkbookGlobalsSink, pivot_cache_stream_paths};
use super::model::{OpenOptions, Reader};
use super::validation_only::CellDecoding;
use crate::defined_names::DefinedNameSlot;
use crate::encryption::prepare_workbook_stream_with_state;
use crate::error::{Error, Result};
use crate::leniency::{ToleranceLog, ToleranceReport};
use crate::number_format::Formatting;
use crate::pivot_table;
use crate::records::{BiffVersion, Encoding};
use crate::sheet_metadata::SheetMetadata;
use crate::{LimitKind, SourceLimits};
use litchi_biff::Records;
use litchi_cfb::{OleFile, OleFileLimits};
use litchi_core::edit::RefusalKind;
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::io::{Read, Seek, SeekFrom};
use std::sync::Arc;

impl<R: Read + Seek> Reader<R> {
    pub(super) fn empty(ole_file: OleFile<R>, source_limits: SourceLimits) -> Self {
        Self {
            ole_file,
            source_limits,
            encryption: None,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            sheets: Arc::from(Vec::<SheetMetadata>::new()),
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
            xml_map: None,
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            formatting: Arc::new(Formatting::default()),
            protection: crate::protection::WorkbookProtection::default(),
            calculation: crate::calculation::WorkbookCalculation::default(),
            picture_compression: None,
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

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(reader: R) -> Result<Self> {
        Self::new_with_options(reader, OpenOptions::default())
    }

    /// Open an XLS workbook with an explicit password contract.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new_with_options(reader: R, options: OpenOptions<'_>) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let ole_file = Self::open_ole_file_with_limits(reader, source_limits)?;
        preflight_declared_streams(&ole_file, source_limits)?;
        Self::from_admitted_ole_file(ole_file, options)
    }

    /// Create an XLS workbook from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - An already-parsed OLE file
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_ole_file(ole_file: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_options(ole_file, OpenOptions::default())
    }

    /// Create a workbook from a parsed OLE file with explicit open options.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_ole_file_with_options(
        ole_file: OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let source_limits = options.source_limits_value();
        Self::validate_ole_file(&ole_file, source_limits)?;
        Self::from_admitted_ole_file(ole_file, options)
    }

    /// Decode a CFB owner whose physical admission is already complete.
    /// The positional materializer uses its existing source/materialization
    /// bounds, while public eager entry points use `SourceLimits` above.
    pub(crate) fn from_admitted_ole_file(
        ole_file: OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<Self> {
        let mut workbook = Self::empty(ole_file, options.source_limits_value());
        workbook.xml_map = crate::xml_map::parse_stream_if_present(&mut workbook.ole_file)?;
        workbook.parse_workbook(&options, CellDecoding::Every)?;
        Ok(workbook)
    }

    /// Create a workbook from a parsed OLE file, or hand the file back when
    /// the workbook needs a password to open.
    ///
    /// This is [`Self::from_ole_file_with_options`] for callers that classify
    /// encrypted input without discarding it: when the `FILEPASS` record asks
    /// for a password that `options` does not supply, the result is
    /// `Ok(Err(ole_file))`, and the container can be opened again with a
    /// password without parsing it a second time.
    ///
    /// # Errors
    ///
    /// Returns every error of [`Self::from_ole_file_with_options`] other than
    /// [`Error::PasswordRequired`].
    pub fn from_ole_file_or_password_required(
        ole_file: OleFile<R>,
        options: OpenOptions<'_>,
    ) -> Result<std::result::Result<Self, OleFile<R>>> {
        let source_limits = options.source_limits_value();
        Self::validate_ole_file(&ole_file, source_limits)?;
        let mut workbook = Self::empty(ole_file, source_limits);

        workbook.xml_map = crate::xml_map::parse_stream_if_present(&mut workbook.ole_file)?;
        match workbook.parse_workbook(&options, CellDecoding::Every) {
            Ok(()) => Ok(Ok(workbook)),
            Err(Error::PasswordRequired) => Ok(Err(workbook.ole_file)),
            Err(error) => Err(error),
        }
    }

    /// Open a CFB source after checking the complete physical input size and
    /// applying the corresponding low-level CFB input ceiling.
    pub(super) fn open_ole_file_with_limits(
        mut reader: R,
        source_limits: SourceLimits,
    ) -> Result<OleFile<R>> {
        preflight_source_len(&mut reader, source_limits)?;
        let cfb_limit = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
        Ok(OleFile::open(reader, OleFileLimits::new(cfb_limit)?)?)
    }

    /// Validate a parsed CFB source under the XLS physical ingress policy.
    ///
    /// Adoption cannot constrain allocations that occurred while the caller
    /// parsed the OLE file. It does ensure that no XML, Workbook, or other
    /// declared stream payload is read before the selected policy is checked.
    pub(super) fn validate_ole_file(
        ole_file: &OleFile<R>,
        source_limits: SourceLimits,
    ) -> Result<()> {
        preflight_ole_file(ole_file, source_limits)
    }

    /// Parse the workbook stream.
    ///
    /// `cells` selects which worksheet cells are decoded. It changes nothing
    /// else: every record is validated by the same code in the same order.
    pub(super) fn parse_workbook(
        &mut self,
        options: &OpenOptions<'_>,
        cells: CellDecoding<'_>,
    ) -> Result<()> {
        // Find and read the Workbook stream
        let workbook_data = self
            .ole_file
            .open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;
        let (workbook_data, encryption) =
            prepare_workbook_stream_with_state(workbook_data, options.password_value())?;
        self.encryption = encryption;
        let mut tolerance = ToleranceLog::new(options.leniency_value());
        // Container repairs come first: the compound-file reader applied them
        // while the OLE file was opened, before any BIFF record was read.
        tolerance
            .report_mut()
            .append_container(self.ole_file.tolerance_report());

        let mut records = Records::new(&workbook_data, litchi_biff::Limits::default())?;
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

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.shared_string_properties = Some(Arc::new(string_properties));
        let all_sheet_names = bound_sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        crate::records::check_duplicate_sheet_names(
            all_sheet_names.iter().map(String::as_str),
            &mut tolerance,
        )?;
        self.sheets = Arc::from(
            bound_sheets
                .iter()
                .enumerate()
                .map(|(index, sheet)| SheetMetadata::from_bound_sheet(index, sheet))
                .collect::<Vec<_>>(),
        );
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
            match self.parse_worksheet_decoding(
                &workbook_data,
                sheet_index,
                bound_sheet,
                &encoding,
                options.compatibility_profile_value(),
                cells,
                tolerance.report_mut(),
            ) {
                Ok(worksheet) => {
                    let worksheet_index = self.worksheets.len();
                    self.worksheet_names.push(bound_sheet.name.clone());
                    self.worksheets.push(worksheet);
                    Arc::make_mut(&mut self.sheets)[sheet_index]
                        .set_parsed_worksheet_index(worksheet_index);
                },
                Err(error @ Error::InvalidRecord { record_type, .. })
                    if pivot_table::is_worksheet_view_record(record_type) =>
                {
                    return Err(error);
                },
                Err(_) => {},
            }
        }
        self.tolerance = tolerance.into_report();
        // The metadata streams' repairs follow every BIFF-level repair.
        let ole_file = &mut self.ole_file;
        self.tolerance
            .append_property_sets(|path| ole_file.property_set_stream(path));

        crate::xml_map::validate_list_columns(self.xml_map.as_ref(), &self.worksheets)?;

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
                    .map(crate::view::View::is_selected)
            })
            .collect::<Vec<_>>();
        self.workbook_view
            .validate_sheet_state(&visible_tabs, &selected_worksheet_tabs)?;

        Ok(())
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn custom_xml_data_store(
        &mut self,
        limits: litchi_ole_common::custom_xml::Limits,
    ) -> Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole_file, limits).map_err(Error::from)
    }

    /// Source limits inherited by package-stream readers.
    #[must_use]
    pub const fn source_limits(&self) -> SourceLimits {
        self.source_limits
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify workbook XML signatures with an explicit trust-neutral policy.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn signatures(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole_file, litchi_sign::cfb::Format::Xls, policy)
            .map_err(Error::from)
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn document_summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(litchi_cfb::OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Read the optional Office Toolbars (`XCB`) stream as inert typed metadata.
    ///
    /// The returned value owns decoded strings and visual bytes because the
    /// compound-file stream buffer is temporary. No command, macro, UI, or
    /// `ActiveX` behavior is activated while reading this stream.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn toolbar(&mut self) -> Result<Option<crate::Wrapper<'static>>> {
        match self.ole_file.open_stream(&["XCB"]) {
            Ok(data) => crate::toolbar::parse(&data)
                .map(|value| Some(value.into_owned()))
                .map_err(|error| Error::InvalidData(format!("XCB toolbar: {error}"))),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
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

    /// Parse the `_VBA_PROJECT_CUR` MS-OVBA project with explicit limits.
    ///
    /// The method returns `None` when no structurally complete VBA project is
    /// present. Source is only decompressed and decoded; it is never compiled,
    /// interpreted, or executed.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn vba(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> Result<Option<litchi_vba::project::Project>> {
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
        litchi_vba::project::Project::open(&mut self.ole_file, &path, limits)
            .map(Some)
            .map_err(Error::from)
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
    /// `limits` is required at this raw boundary. It is constrained by the
    /// source profile retained when this reader was admitted, so a caller can
    /// narrow the operation but cannot widen the ingress policy.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn revision_log(
        &mut self,
        limits: &SourceLimits,
    ) -> Result<Option<crate::revision_log::RevisionLog>> {
        let Some(name) =
            crate::revision_log::find_revision_log_stream(&self.ole_file.list_streams())
                .map(str::to_owned)
        else {
            return Ok(None);
        };
        let effective_limits = self.source_limits.constrained_by(*limits);
        let maximum = effective_limits.max_entry_bytes();
        let observed = self.ole_file.stream_len(&[name.as_str()])?;
        let maximum_u64 = u64::try_from(maximum).unwrap_or(u64::MAX);
        if observed > maximum_u64 {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed,
                maximum: maximum_u64,
            });
        }
        let data = self.ole_file.open_stream(&[name.as_str()])?;
        crate::revision_log::parse_revision_log_stream_with_sheet_count(&data, self.sheets.len())
            .map(Some)
    }

    /// Whether this workbook contains the shared-workbook `User Names` stream.
    ///
    /// The stream is intentionally discovered without opening or parsing it;
    /// [`Self::user_names`] performs the bounded lazy read.
    #[must_use]
    pub fn has_user_names(&self) -> bool {
        self.ole_file.list_streams().iter().any(|path| {
            path.len() == 1
                && path[0].eq_ignore_ascii_case(crate::user_names::USER_NAMES_STREAM_NAME)
        })
    }

    /// Parse the shared-workbook `User Names` stream with explicit resource
    /// bounds. The read remains lazy and inert, and the Revision Log source
    /// identity is retained for source-checked patch application.
    /// # Errors
    ///
    /// Returns an error when a bound, stream grammar, required Revision Log,
    /// or GUID dependency closure is invalid.
    pub fn user_names(
        &mut self,
        limits: crate::user_names::Limits,
    ) -> Result<Option<crate::user_names::Snapshot>> {
        let limits = limits.validate()?;
        let streams = self.ole_file.list_streams();
        let names = streams
            .iter()
            .filter(|path| {
                path.len() == 1
                    && path[0].eq_ignore_ascii_case(crate::user_names::USER_NAMES_STREAM_NAME)
            })
            .map(|path| path[0].clone())
            .collect::<Vec<_>>();
        if names.len() > 1 {
            return Err(Error::InvalidData(
                "CFB contains more than one User Names stream".to_string(),
            ));
        }
        let Some(name) = names.first() else {
            return Ok(None);
        };
        let max_stream_bytes = u64::try_from(limits.max_stream_bytes).unwrap_or(u64::MAX);

        let revision_names = streams
            .iter()
            .filter(|path| {
                path.len() == 1
                    && path[0].eq_ignore_ascii_case(crate::revision_log::REVISION_LOG_STREAM_NAME)
            })
            .map(|path| path[0].clone())
            .collect::<Vec<_>>();
        if revision_names.len() > 1 {
            return Err(Error::InvalidData(
                "CFB contains more than one Revision Log stream".to_string(),
            ));
        }
        let Some(revision_name) = revision_names.first() else {
            return Err(Error::InvalidData(
                "User Names stream requires a Revision Log stream".to_string(),
            ));
        };
        let revision_len = self.ole_file.stream_len(&[revision_name.as_str()])?;
        if revision_len > max_stream_bytes {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit workbook",
                "the Revision Log stream exceeds its byte limit",
            ));
        }

        let stream_len = self.ole_file.stream_len(&[name.as_str()])?;
        if stream_len > max_stream_bytes {
            return Err(Error::refused(
                RefusalKind::UnsupportedOperation,
                "edit workbook",
                "the User Names stream exceeds its byte limit",
            ));
        }
        let data = Arc::<[u8]>::from(
            self.ole_file
                .open_stream(&[name.as_str()])?
                .into_boxed_slice(),
        );
        let revision_data = Arc::<[u8]>::from(
            self.ole_file
                .open_stream(&[revision_name.as_str()])?
                .into_boxed_slice(),
        );
        let revision_guids = crate::revision_log::scan_revision_header_guids_with_sheet_count(
            &revision_data,
            limits.max_stream_bytes,
            limits.max_revision_records,
            limits.max_revision_guids,
            self.sheets.len(),
        )?;
        crate::user_names::Snapshot::parse_shared_with_revision_guids_arc_and_revision_log_with_limits(
            data,
            revision_guids,
            revision_data,
            limits,
        )
        .map(Some)
    }
}

/// Validate the physical size and declared stream lengths of an already
/// parsed CFB container before any format payload is materialized.
pub(crate) fn preflight_ole_file<R: Read + Seek>(
    ole_file: &OleFile<R>,
    source_limits: SourceLimits,
) -> Result<()> {
    validate_source_len(ole_file.file_size(), source_limits)?;
    preflight_declared_streams(ole_file, source_limits)?;
    Ok(())
}

/// Check the complete source length without allocating or entering the CFB
/// parser. A `Read + Seek` source normally reports its length directly; the
/// bounded probe is retained for readers whose seek-to-end operation fails.
fn preflight_source_len<R: Read + Seek>(reader: &mut R, source_limits: SourceLimits) -> Result<()> {
    let original_position = reader.stream_position()?;
    match reader.seek(SeekFrom::End(0)) {
        Ok(source_len) => {
            reader.seek(SeekFrom::Start(original_position))?;
            validate_source_len(source_len, source_limits)
        },
        Err(_seek_error) => {
            reader.seek(SeekFrom::Start(0))?;
            let maximum = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
            let probe_len = maximum.saturating_add(1);
            let mut observed = 0_u64;
            let mut buffer = [0_u8; 8192];
            let count_result = (|| -> std::io::Result<()> {
                while observed < probe_len {
                    let remaining = probe_len - observed;
                    let request = usize::try_from(remaining)
                        .unwrap_or(usize::MAX)
                        .min(buffer.len());
                    let read = reader.read(&mut buffer[..request])?;
                    if read == 0 {
                        break;
                    }
                    observed = observed.saturating_add(u64::try_from(read).unwrap_or(u64::MAX));
                }
                Ok(())
            })();
            let restore_result = reader.seek(SeekFrom::Start(original_position));
            count_result?;
            restore_result?;
            validate_source_len(observed, source_limits)
        },
    }
}

fn validate_source_len(actual: u64, source_limits: SourceLimits) -> Result<()> {
    let maximum = u64::try_from(source_limits.max_input_bytes()).unwrap_or(u64::MAX);
    if actual > maximum {
        return Err(Error::Limit {
            kind: LimitKind::InputBytes,
            observed: actual,
            maximum,
        });
    }
    Ok(())
}

/// Check all declared CFB stream lengths without reading any stream payload.
///
/// The aggregate includes opaque, metadata, and embedded-object streams. The
/// root mini-stream backing is excluded because charging it in addition to
/// its stream entries would count mini-stream payloads twice.
fn preflight_declared_streams<R: Read + Seek>(
    ole_file: &OleFile<R>,
    source_limits: SourceLimits,
) -> Result<u64> {
    let entry_maximum = u64::try_from(source_limits.max_entry_bytes()).unwrap_or(u64::MAX);
    let aggregate_maximum = u64::try_from(source_limits.max_aggregate_bytes()).unwrap_or(u64::MAX);
    let mut aggregate = 0_u64;
    for index in 0..ole_file.directory_entry_count() {
        let sid = u32::try_from(index).map_err(|_error| {
            Error::InvalidData("XLS CFB directory entry index exceeds u32".to_string())
        })?;
        let Some(entry) = ole_file.directory_entry_by_sid(sid) else {
            continue;
        };
        if entry.entry_type != litchi_cfb::consts::STGTY_STREAM {
            continue;
        }
        if entry.size > entry_maximum {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: entry.size,
                maximum: entry_maximum,
            });
        }
        aggregate = aggregate.checked_add(entry.size).ok_or(Error::Limit {
            kind: LimitKind::AggregateBytes,
            observed: u64::MAX,
            maximum: aggregate_maximum,
        })?;
    }
    if aggregate > aggregate_maximum {
        return Err(Error::Limit {
            kind: LimitKind::AggregateBytes,
            observed: aggregate,
            maximum: aggregate_maximum,
        });
    }
    Ok(aggregate)
}

#[cfg(test)]
mod source_limit_tests {
    use super::{preflight_declared_streams, preflight_ole_file, preflight_source_len};
    use crate::{Error, LimitKind, OpenOptions, SourceLimits, raw::Reader};
    use litchi_cfb::{OleFile, OleWriter};
    use std::io::{Cursor, Read, Seek, SeekFrom};
    use std::sync::{Arc, atomic::AtomicUsize, atomic::Ordering};

    fn limits(input: usize, entry: usize, aggregate: usize) -> SourceLimits {
        SourceLimits::builder()
            .max_input_bytes(input)
            .unwrap()
            .max_entry_bytes(entry)
            .unwrap()
            .max_aggregate_bytes(aggregate)
            .unwrap()
            .build()
            .unwrap()
    }

    fn cfb_with_streams() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Workbook"], b"abc").unwrap();
        writer.create_stream(&["OpaqueMetadata"], b"12345").unwrap();
        writer.create_stream(&["Nested", "Payload"], b"xy").unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    #[test]
    fn exact_entry_and_aggregate_boundaries_include_opaque_streams() {
        let bytes = cfb_with_streams();
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let source_limits = limits(bytes.len(), 5, 10);
        assert_eq!(preflight_declared_streams(&ole, source_limits).unwrap(), 10);
        preflight_ole_file(&ole, source_limits).unwrap();
    }

    #[test]
    fn entry_limit_precedes_aggregate_limit() {
        let bytes = cfb_with_streams();
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let error = preflight_declared_streams(&ole, limits(bytes.len(), 4, 1)).unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: 5,
                maximum: 4,
            }
        ));
    }

    #[test]
    fn aggregate_limit_accounts_for_opaque_and_nested_streams() {
        let bytes = cfb_with_streams();
        let ole = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let error = preflight_declared_streams(&ole, limits(bytes.len(), 5, 9)).unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::AggregateBytes,
                observed: 10,
                maximum: 9,
            }
        ));
    }

    #[test]
    fn input_limit_is_checked_before_cfb_payload_or_structure() {
        let bytes = cfb_with_streams();
        let error = preflight_source_len(
            &mut Cursor::new(bytes.clone()),
            limits(
                bytes.len() - 1,
                SourceLimits::MAX_ENTRY_BYTES,
                SourceLimits::MAX_AGGREGATE_BYTES,
            ),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::InputBytes,
                observed,
                maximum,
            } if observed == bytes.len() as u64 && maximum == (bytes.len() - 1) as u64
        ));
    }

    #[derive(Debug)]
    struct CountingReader {
        cursor: Cursor<Vec<u8>>,
        reads: Arc<AtomicUsize>,
    }

    impl Read for CountingReader {
        fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
            let count = self.cursor.read(output)?;
            self.reads.fetch_add(count, Ordering::Relaxed);
            Ok(count)
        }
    }

    impl Seek for CountingReader {
        fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
            self.cursor.seek(position)
        }
    }

    #[test]
    fn eager_input_rejection_happens_before_container_reads() {
        let bytes = cfb_with_streams();
        let reads = Arc::new(AtomicUsize::new(0));
        let reader = CountingReader {
            cursor: Cursor::new(bytes.clone()),
            reads: Arc::clone(&reads),
        };
        let source_limits = limits(
            bytes.len() - 1,
            SourceLimits::MAX_ENTRY_BYTES,
            SourceLimits::MAX_AGGREGATE_BYTES,
        );
        let error =
            Reader::new_with_options(reader, OpenOptions::new().source_limits(source_limits))
                .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::InputBytes,
                ..
            }
        ));
        assert_eq!(reads.load(Ordering::Relaxed), 0);
    }

    #[test]
    fn adopted_container_rejects_stream_lengths_before_payload_reads() {
        let bytes = cfb_with_streams();
        let reads = Arc::new(AtomicUsize::new(0));
        let reader = CountingReader {
            cursor: Cursor::new(bytes.clone()),
            reads: Arc::clone(&reads),
        };
        let ole = OleFile::open(reader, litchi_cfb::OleFileLimits::default()).unwrap();
        let reads_after_cfb = reads.load(Ordering::Relaxed);
        let source_limits = limits(bytes.len(), 4, SourceLimits::MAX_AGGREGATE_BYTES);
        let error = Reader::from_ole_file_with_options(
            ole,
            OpenOptions::new().source_limits(source_limits),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: 5,
                maximum: 4,
            }
        ));
        assert_eq!(reads.load(Ordering::Relaxed), reads_after_cfb);
    }

    #[test]
    fn password_returning_adoption_applies_the_same_preflight_gate() {
        let bytes = cfb_with_streams();
        let reads = Arc::new(AtomicUsize::new(0));
        let reader = CountingReader {
            cursor: Cursor::new(bytes.clone()),
            reads: Arc::clone(&reads),
        };
        let ole = OleFile::open(reader, litchi_cfb::OleFileLimits::default()).unwrap();
        let reads_after_cfb = reads.load(Ordering::Relaxed);
        let source_limits = limits(bytes.len(), 4, SourceLimits::MAX_AGGREGATE_BYTES);
        let error = Reader::from_ole_file_or_password_required(
            ole,
            OpenOptions::new().source_limits(source_limits),
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: 5,
                maximum: 4,
            }
        ));
        assert_eq!(reads.load(Ordering::Relaxed), reads_after_cfb);
    }
    struct NoEndReader {
        cursor: Cursor<Vec<u8>>,
        fail_read: bool,
    }

    impl Read for NoEndReader {
        fn read(&mut self, output: &mut [u8]) -> std::io::Result<usize> {
            if self.fail_read {
                self.cursor.set_position(1);
                return Err(std::io::Error::other("injected read failure"));
            }
            self.cursor.read(output)
        }
    }

    impl Seek for NoEndReader {
        fn seek(&mut self, position: SeekFrom) -> std::io::Result<u64> {
            if matches!(position, SeekFrom::End(_)) {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::Unsupported,
                    "end seeking unavailable",
                ));
            }
            self.cursor.seek(position)
        }
    }

    #[test]
    fn bounded_length_probe_restores_position_on_success_limit_and_read_error() {
        for (maximum, fail_read) in [(10, false), (9, false), (10, true)] {
            let mut reader = NoEndReader {
                cursor: Cursor::new(vec![0; 10]),
                fail_read,
            };
            reader.cursor.set_position(3);
            let result = preflight_source_len(&mut reader, limits(maximum, 10, 10));
            assert_eq!(reader.cursor.position(), 3);
            match (maximum, fail_read) {
                (_, true) => assert!(matches!(result, Err(Error::Io(_)))),
                (9, false) => assert!(matches!(
                    result,
                    Err(Error::Limit {
                        kind: LimitKind::InputBytes,
                        observed: 10,
                        maximum: 9,
                    })
                )),
                _ => result.unwrap(),
            }
        }
    }
}
