use super::{Error, OpenOptions, Reader, Result, SourceLimits};
use crate::presentation::Presentation;
use litchi_cfb::{OleError, OleFile};
use litchi_ole_common::property_set::{
    PropertySetReader, Section, Stream, USER_DEFINED_PROPERTIES_FMTID,
};
use std::fs::File;
use std::io::{Read, Seek, SeekFrom};
use std::path::Path;

const POWERPOINT_DOCUMENT_STREAM: &[&str] = &["PowerPoint Document"];

impl Reader<File> {
    /// Open a .ppt package from a file path.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the .ppt file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ppt::raw::Reader;
    ///
    /// let mut pkg = Reader::open("presentation.ppt")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_options(path, OpenOptions::default())
    }

    /// Open a `.ppt` package under explicit open options.
    ///
    /// The reader keeps the options' source limits, leniency, and VBA
    /// inspection profile; every projection it returns ([`Self::presentation`],
    /// the font and picture-bullet readers) inherits them.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn open_with_options<P: AsRef<Path>>(path: P, options: OpenOptions<'_>) -> Result<Self> {
        let file = File::open(path)?;
        Reader::from_reader_with_options(file, options)
    }
}

impl<R: Read + Seek> Reader<R> {
    /// Create a Reader from any source that implements Read + Seek.
    ///
    /// # Arguments
    ///
    /// * `reader` - A reader containing the .ppt file data
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use std::fs::File;
    /// use litchi_ppt::raw::Reader;
    ///
    /// let file = File::open("presentation.ppt")?;
    /// let pkg = Reader::from_reader(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, OpenOptions::default())
    }

    /// Create a reader under explicit open options.
    ///
    /// The reader keeps the options' source limits, leniency, and VBA
    /// inspection profile; every projection it returns inherits them.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_reader_with_options(mut reader: R, options: OpenOptions<'_>) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let original_position = reader.stream_position()?;
        let input_len = reader.seek(SeekFrom::End(0))?;
        reader.seek(SeekFrom::Start(original_position))?;
        let input_bytes = usize::try_from(input_len).map_err(|_err| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if input_bytes > source_limits.max_package_bytes {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: input_bytes,
                maximum: source_limits.max_package_bytes,
            });
        }
        let ole = OleFile::open(reader, litchi_cfb::OleFileLimits::default())?;

        preflight_declared_streams(&ole, source_limits)?;

        validate_powerpoint_document_stream(&ole)?;

        Ok(Self {
            ole,
            source_limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
        })
    }

    /// Create a Reader from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole` - An already-parsed OLE file
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_cfb::{OleFile, OleFileLimits};
    /// use litchi_ppt::raw::Reader;
    /// use std::io::Cursor;
    ///
    /// let bytes = std::fs::read("presentation.ppt")?;
    /// let ole = OleFile::open(Cursor::new(bytes), OleFileLimits::default())?;
    /// let pkg = Reader::from_ole_file(ole)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_ole_file(ole: OleFile<R>) -> Result<Self> {
        Self::from_ole_file_with_options(ole, OpenOptions::default())
    }

    /// Wrap an already-parsed OLE file under explicit open options.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_ole_file_with_options(ole: OleFile<R>, options: OpenOptions<'_>) -> Result<Self> {
        let source_limits = options.source_limits_value();
        let input_bytes = usize::try_from(ole.file_size()).map_err(|_error| Error::Limit {
            kind: super::SourceLimitKind::InputBytes,
            observed: usize::MAX,
            maximum: source_limits.max_package_bytes,
        })?;
        if input_bytes > source_limits.max_package_bytes {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::InputBytes,
                observed: input_bytes,
                maximum: source_limits.max_package_bytes,
            });
        }
        preflight_declared_streams(&ole, source_limits)?;
        validate_powerpoint_document_stream(&ole)?;

        Ok(Self {
            ole,
            source_limits,
            leniency: options.leniency_value(),
            #[cfg(feature = "vba-inspection")]
            vba_project_limits: options.vba_project_limits_value(),
        })
    }

    /// Get the main presentation.
    ///
    /// Returns the `Presentation` object which provides access to the presentation's
    /// content, slides, and other features.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ppt::raw::Reader;
    ///
    /// let mut pkg = Reader::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation(&mut self) -> Result<Presentation> {
        let options = OpenOptions::default().leniency(self.leniency);
        #[cfg(feature = "vba-inspection")]
        let options = options.vba_project_limits(self.vba_project_limits);
        Presentation::from_ole_with_options(&mut self.ole, options, self.source_limits)
    }

    /// Get the main presentation with per-call open options, such as the
    /// password of an encrypted presentation.
    ///
    /// Explicit source limits in `options` can only tighten the reader's
    /// own; the options' leniency replaces the reader's for this call.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation_with_options(&mut self, options: OpenOptions<'_>) -> Result<Presentation> {
        #[cfg(feature = "vba-inspection")]
        let options = options.inherit_vba_project_limits(self.vba_project_limits);
        let source_limits = if options.has_explicit_source_limits() {
            self.source_limits
                .constrained_by(options.source_limits_value())
        } else {
            self.source_limits
        };
        Presentation::from_ole_with_options(&mut self.ole, options, source_limits)
    }

    /// Limits inherited by [`Self::presentation`].
    pub fn source_limits(&self) -> SourceLimits {
        self.source_limits
    }

    /// Leniency inherited by [`Self::presentation`].
    pub fn leniency(&self) -> crate::leniency::Leniency {
        self.leniency
    }

    /// Whether the already-parsed CFB directory contains a digital-signature
    /// stream or storage, including one nested below a storage.  This is kept
    /// on the reader so source admission does not parse the same CFB a second
    /// time merely to make the signed/no-op mutation decision.
    pub(crate) fn contains_signature_component(&self) -> Result<bool> {
        let max_entries = self
            .source_limits
            .max_events()
            .min(self.ole.directory_entry_count());
        let mut pending = vec![Vec::<String>::new()];
        let mut discovered = 0usize;
        while let Some(path) = pending.pop() {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let mut found = false;
            let remaining = max_entries.saturating_sub(discovered);
            self.ole
                .visit_directory_entry_refs(&refs, remaining, |entry| {
                    discovered = discovered.saturating_add(1);
                    if is_signature_component(&entry.name) {
                        found = true;
                    }
                    if entry.entry_type == 1 {
                        let mut child = path.clone();
                        child.push(entry.name.clone());
                        pending.push(child);
                    }
                    Ok::<(), OleError>(())
                })?;
            if found {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Read the live document-comparison snapshot from the presentation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn document_comparison(&mut self) -> Result<crate::document_comparison::Snapshot> {
        self.presentation()?.document_comparison()
    }

    /// Get the underlying OLE file.
    ///
    /// This provides access to lower-level OLE operations and streams.
    #[inline]
    pub fn ole_file(&mut self) -> &mut OleFile<R> {
        &mut self.ole
    }

    /// Consumes the package and returns its parsed OLE file.
    ///
    /// The parsed container index is kept, so a caller that has finished
    /// reading the presentation can hand the same parse to another owner (for
    /// example `litchi_cfb::SharedOleFile::from_owned_ole_file`) instead of
    /// parsing the container again.
    #[must_use]
    pub fn into_ole_file(self) -> OleFile<R> {
        self.ole
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn custom_xml_data_store(
        &mut self,
        limits: litchi_ole_common::custom_xml::Limits,
    ) -> Result<Option<litchi_ole_common::custom_xml::Store>> {
        litchi_ole_common::custom_xml::inspect(&mut self.ole, limits).map_err(Error::from)
    }

    /// Read the legacy Summary Information property set stream.
    ///
    /// Returns `None` when the document does not contain the stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
    pub fn summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify presentation XML signatures with an explicit trust-neutral policy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "sign")]
    pub fn signatures(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole, litchi_sign::cfb::Format::Ppt, policy)
            .map_err(Error::from)
    }

    /// Read the legacy Document Summary Information property set stream.
    ///
    /// Returns `None` when the document does not contain the stream.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
    pub fn document_summary_information(&mut self) -> Result<Option<Stream>> {
        match self
            .ole
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_ole_common::Error::Cfb(OleError::StreamNotFound)) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Read the user-defined properties section of the Document Summary
    /// Information property set.
    ///
    /// Returns `None` when the document has no Document Summary Information
    /// stream or no user-defined properties section.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying OLE container cannot be read.
    pub fn user_defined_properties(&mut self) -> Result<Option<Section>> {
        Ok(self
            .document_summary_information()?
            .and_then(|stream| stream.section(USER_DEFINED_PROPERTIES_FMTID).cloned()))
    }
}

pub(crate) fn is_signature_component(name: &str) -> bool {
    [
        "_xmlsignatures",
        "_signatures",
        "DigitalSignature",
        "\u{0005}DigitalSignature",
    ]
    .iter()
    .any(|marker| name.eq_ignore_ascii_case(marker))
}

fn validate_powerpoint_document_stream<R: Read + Seek>(ole: &OleFile<R>) -> Result<()> {
    if ole.exists(POWERPOINT_DOCUMENT_STREAM) && !ole.directory_exists(POWERPOINT_DOCUMENT_STREAM) {
        return Ok(());
    }

    Err(Error::InvalidFormat(
        "Not a valid PowerPoint document: PowerPoint Document stream not found or is not a stream"
            .to_string(),
    ))
}

/// Check every declared CFB stream before a PPT owner exposes APIs that can
/// materialize arbitrary metadata or embedded streams. The root mini-stream
/// backing is excluded because its bytes are charged through the child stream
/// entries, matching the XLS source policy.
pub(super) fn preflight_declared_streams<R: Read + Seek>(
    ole: &OleFile<R>,
    source_limits: SourceLimits,
) -> Result<u64> {
    let entries = (0..ole.directory_entry_count()).filter_map(|index| {
        let sid = u32::try_from(index).ok()?;
        ole.directory_entry_by_sid(sid)
    });
    preflight_directory_entries(entries, source_limits)
}

pub(super) fn preflight_directory_entries<'a>(
    entries: impl Iterator<Item = &'a litchi_cfb::DirectoryEntry>,
    source_limits: SourceLimits,
) -> Result<u64> {
    let entry_maximum = source_limits.max_input_bytes;
    let aggregate_maximum = source_limits.max_aggregate_input_bytes;
    let aggregate_maximum = u64::try_from(aggregate_maximum).unwrap_or(u64::MAX);
    let mut aggregate = 0_u64;
    for entry in entries {
        if entry.entry_type != litchi_cfb::consts::STGTY_STREAM {
            continue;
        }
        let size = usize::try_from(entry.size).map_err(|_error| Error::Limit {
            kind: super::SourceLimitKind::EntryBytes,
            observed: usize::MAX,
            maximum: entry_maximum,
        })?;
        if size > entry_maximum {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::EntryBytes,
                observed: size,
                maximum: entry_maximum,
            });
        }
        aggregate = aggregate.checked_add(entry.size).ok_or(Error::Limit {
            kind: super::SourceLimitKind::AggregateBytes,
            observed: usize::MAX,
            maximum: usize::try_from(aggregate_maximum).unwrap_or(usize::MAX),
        })?;
        if aggregate > aggregate_maximum {
            return Err(Error::Limit {
                kind: super::SourceLimitKind::AggregateBytes,
                observed: usize::try_from(aggregate).unwrap_or(usize::MAX),
                maximum: usize::try_from(aggregate_maximum).unwrap_or(usize::MAX),
            });
        }
    }
    Ok(aggregate)
}
