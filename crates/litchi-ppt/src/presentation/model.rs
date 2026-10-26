use crate::consts::RecordType;
use crate::document_properties::DocumentProperties12;
use crate::embedded::object::Collection as OleCollection;
#[cfg(feature = "vba-inspection")]
use crate::embedded::storage::Ref as StorageRef;
use crate::embedded::storage::{Kind as StorageKind, Storage};
use crate::external_media::{Collection as MediaCollection, Limits as MediaLimits};
use crate::header_footer::{
    HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope, HeaderFooters,
};
use crate::hyperlink::Hyperlinks;
use crate::main_master::MainMasterMetadata12;
use crate::non_zoom_view::OutlineSorterViewInformation;
/// High-performance Presentation API with zero-copy slide parsing.
use crate::package::{Error, Result, SourceLimits};
use crate::parsers::RecordParser;
use crate::persist::PersistMapping;
use crate::records::Record;
use crate::routing_slip::Slip;
use crate::slide::{ParsedComment, Slide, SlideDirectory, SlideFactory};
use crate::sound_collection::Collection;
use crate::view_info::SlideViewInformation;
#[cfg(feature = "vba-inspection")]
use litchi_cfb::OleFile;
use litchi_cfb::SharedOleFile;
use litchi_odraw::image::{File as ImageFile, Id as ImageId, Store as ImageStore};
#[cfg(feature = "vba-inspection")]
use std::io::Cursor;
use std::sync::{Arc, OnceLock};

#[cfg(feature = "encryption")]
use crate::encryption::EncryptionProfile;

/// Deferred native `Pictures` stream owned by a positional presentation.
///
/// The descriptor is created from the validated CFB directory and is kept
/// independently from the payload cache.  A successful materialization is
/// published exactly once; failed reads are deliberately left retryable.
pub(super) struct LazyPictures {
    shared: Arc<SharedOleFile>,
    path: &'static [&'static str],
    length: usize,
    limits: SourceLimits,
    #[cfg(feature = "encryption")]
    crypto: Option<litchi_crypto::rc4::Context>,
    cache: OnceLock<Arc<[u8]>>,
}

impl LazyPictures {
    pub(super) fn new(
        shared: Arc<SharedOleFile>,
        path: &'static [&'static str],
        length: usize,
        limits: SourceLimits,
        #[cfg(feature = "encryption")] crypto: Option<litchi_crypto::rc4::Context>,
    ) -> Self {
        Self {
            shared,
            path,
            length,
            limits,
            #[cfg(feature = "encryption")]
            crypto,
            cache: OnceLock::new(),
        }
    }

    pub(super) fn bytes(&self) -> Result<&[u8]> {
        // This check is intentionally before the cache branch: a stale source
        // must not make a previously successful payload appear current.
        self.shared.source_version()?;
        if self.cache.get().is_none() {
            if self.length > self.limits.max_input_bytes {
                return Err(Error::Limit {
                    kind: crate::SourceLimitKind::EntryBytes,
                    observed: self.length,
                    maximum: self.limits.max_input_bytes,
                });
            }
            let declared = self.shared.stream_len(self.path)?;
            let declared = usize::try_from(declared).map_err(|_error| Error::Limit {
                kind: crate::SourceLimitKind::EntryBytes,
                observed: usize::MAX,
                maximum: self.limits.max_input_bytes,
            })?;
            if declared != self.length {
                return Err(Error::Corrupted(
                    "Pictures stream length changed after positional open".to_string(),
                ));
            }
            #[cfg(feature = "encryption")]
            let mut data = self.shared.open_stream(self.path)?;
            #[cfg(not(feature = "encryption"))]
            let data = self.shared.open_stream(self.path)?;
            if data.len() != self.length {
                return Err(Error::Corrupted(
                    "Pictures stream was truncated during positional read".to_string(),
                ));
            }
            #[cfg(feature = "encryption")]
            if let Some(crypto) = &self.crypto {
                crate::encryption::decrypt_pictures(&mut data, crypto, self.limits)?;
            }
            // Check again after the potentially expensive decode/decrypt and
            // before publishing the cache entry.
            self.shared.source_version()?;
            let _ = self.cache.set(data.into());
        }
        self.cache
            .get()
            .map(Arc::as_ref)
            .ok_or_else(|| Error::Corrupted("Pictures cache was not published".to_string()))
    }
}

/// A `PowerPoint` presentation (.ppt) with high-performance zero-copy parsing.
///
/// # Performance
///
/// - Document data loaded once and borrowed for all slides
/// - Slides parsed lazily using persist mapping
/// - Shapes loaded on-demand
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ppt::Package;
///
/// let mut pkg = Package::open("presentation.ppt")?;
/// let pres = pkg.presentation()?;
///
/// // Get slides (zero-copy, lazy evaluation)
/// for slide in pres.slides()? {
///     println!("Slide {}: {}", slide.slide_number(), slide.text()?);
///     println!("  Shapes: {}", slide.shape_count()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Clone)]
pub struct Presentation {
    pub(crate) inner: Arc<PresentationData>,
}

/// Immutable state shared by cheap [`Presentation`] handles.
pub(crate) struct PresentationData {
    /// The main document stream data, shared with the record tree whose
    /// payloads span it (owned for lifetime management)
    pub(crate) powerpoint_document: Arc<Vec<u8>>,
    /// Parsed record structure (reserved for future advanced parsing)
    #[allow(dead_code, reason = "reserved for future advanced parsing")]
    pub(crate) parser: RecordParser,
    /// Persist ID to offset mapping
    pub(crate) persist_mapping: PersistMapping,
    pub(crate) slide_directory: SlideDirectory,
    /// Pictures stream data (for image extraction)
    pub(crate) pictures_data: Option<Vec<u8>>,
    /// Positional Pictures descriptor and success-only payload cache.
    pub(super) pictures_source: Option<LazyPictures>,
    /// Record limits retained for lazy live-persist parsing.
    pub(crate) source_limits: SourceLimits,
    /// VBA inspection limits retained from the owning open options.
    #[cfg(feature = "vba-inspection")]
    pub(crate) vba_project_limits: crate::VbaProjectLimits,
    /// Non-structural defects repaired while the presentation was opened.
    ///
    /// Aggregated by the constructors in `package.rs`; see
    /// [`crate::leniency`] for where future recording sites plug in.
    pub(crate) tolerance: crate::leniency::ToleranceReport,
    /// Positional source retained for freshness checks on fallible reads.
    pub(crate) source: Option<Arc<SharedOleFile>>,
    /// Validated password-to-open profile, when the document was encrypted.
    #[cfg(feature = "encryption")]
    pub(crate) encryption: Option<EncryptionProfile>,
    /// The most recently parsed slide's shapes (record 0260).
    pub(crate) shape_memo: crate::slide::ShapeMemo,
}

impl Presentation {
    pub(super) fn from_data(data: PresentationData) -> Self {
        Self {
            inner: Arc::new(data),
        }
    }

    /// Verify that the retained positional source still matches this projection.
    ///
    /// Owned in-memory projections always pass. This also fences cached reads
    /// after the package's edit ownership has been released.
    ///
    /// # Errors
    /// Returns the typed source error if the source changed or cannot be read.
    ///
    /// ```
    /// let package = litchi_ppt::Package::new()?;
    /// let presentation = package.presentation()?;
    /// presentation.check_source()?;
    /// # Ok::<(), litchi_ppt::Error>(())
    /// ```
    pub fn check_source(&self) -> Result<()> {
        if let Some(source) = self.inner.source.as_ref() {
            source.source_version()?;
        }
        Ok(())
    }

    /// Limits retained by this immutable presentation projection.
    #[must_use]
    pub fn source_limits(&self) -> SourceLimits {
        self.inner.source_limits
    }

    /// Limits retained for inert VBA project inspection.
    #[cfg(feature = "vba-inspection")]
    #[must_use]
    pub fn vba_project_limits(&self) -> crate::VbaProjectLimits {
        self.inner.vba_project_limits
    }

    /// Whether this presentation was opened with password-to-open encryption.
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        #[cfg(feature = "encryption")]
        {
            self.inner.encryption.is_some()
        }
        #[cfg(not(feature = "encryption"))]
        {
            false
        }
    }

    /// The validated password-to-open encryption profile, if present.
    #[cfg(feature = "encryption")]
    #[must_use]
    pub fn encryption(&self) -> Option<EncryptionProfile> {
        self.inner.encryption
    }
}

impl Presentation {
    /// Every non-structural defect the reader repaired while opening this
    /// presentation, in discovery order.
    ///
    /// This is the open-time aggregate: the record tree, `UserEdit`/persist
    /// chain, and slide directory are parsed at open, and the chain and
    /// slide-list repairs ([`crate::PersistDefect`],
    /// [`crate::SlideListDefect`]) made there are listed here, after the
    /// compound-file repairs the container reader applied while the OLE
    /// file was opened ([`crate::Defect::Container`]) and before the
    /// property-set repairs made while the standard metadata streams were
    /// read ([`crate::Defect::PropertySet`]).
    /// Text-style and shape repairs ([`crate::TextPropDefect`],
    /// [`crate::ShapeDefect`]) fire while shapes are lazily parsed; they are
    /// reported per slide by
    /// [`Slide::tolerance_report`] and for the whole presentation by
    /// [`Self::collect_tolerance_report`].
    #[must_use]
    pub fn tolerance_report(&self) -> &crate::leniency::ToleranceReport {
        &self.inner.tolerance
    }

    /// Every non-structural defect the reader repairs when reading this whole
    /// presentation: the open-time aggregate ([`Self::tolerance_report`]),
    /// then the repairs made while each main master's drawing is parsed
    /// (by [`Self::header_footers`]), then those made while each slide's
    /// shapes and each slide's speaker-notes shapes are parsed, in slide
    /// order.
    ///
    /// Each entry is located by the persist ID of its slide or notes page and
    /// the shape ID of its text box (see [`crate::DefectLocation`]). An empty
    /// report means the presentation conformed. This walks every slide, so it
    /// parses every slide's and notes page's shapes; the per-slide
    /// [`Slide::tolerance_report`] is the incremental alternative.
    ///
    /// # Errors
    ///
    /// Returns an error if a slide, notes page, or shape tree cannot be
    /// parsed — the structural failures that stay fatal under every
    /// [`crate::Leniency`] variant.
    pub fn collect_tolerance_report(&self) -> Result<crate::leniency::ToleranceReport> {
        self.check_source()?;
        let mut report = self.inner.tolerance.clone();
        // Main-master drawings are parsed by `Self::header_footers`; their
        // repairs are listed here, before the slides'.
        for master in self
            .inner
            .parser
            .find_records_ref()
            .iter()
            .filter(|record| record.record_type == RecordType::MainMaster)
        {
            super::codec::record_drawing_repairs(master, self.inner.source_limits, &mut report)?;
        }
        for slide in self.slides()? {
            slide.shapes_with_report(&mut report)?;
            if let Some(notes) = slide.speaker_notes()? {
                notes.shapes_with_report(&mut report)?;
            }
        }
        Ok(report)
    }

    /// Get iterator over all slides with zero-copy borrowing.
    ///
    /// # Performance
    ///
    /// - Returns lazy iterator (slides parsed on iteration)
    /// - Zero-copy: slides borrow from document data
    /// - Each slide lazily loads its shapes
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slides(&self) -> Result<Vec<Slide<'_>>> {
        self.check_source()?;
        let factory = SlideFactory::new_shared_with_limits(
            &self.inner.powerpoint_document,
            &self.inner.persist_mapping,
            &self.inner.slide_directory,
            self.inner.source_limits,
        );

        factory
            .slides()
            .enumerate()
            .map(|(idx, slide_result)| {
                slide_result.map(|slide_data| {
                    Slide::from_slide_data(slide_data, idx + 1)
                        .with_shape_memo(&self.inner.shape_memo)
                })
            })
            .collect()
    }

    /// Get one slide by its zero-based logical position.
    ///
    /// Unlike [`Self::slides`], this method parses only the requested slide
    /// and does not retain the other slides in a temporary collection.
    pub fn slide_at(&self, position: usize) -> Result<Option<Slide<'_>>> {
        self.check_source()?;
        let Some(entry) = self.inner.slide_directory.entries().get(position) else {
            return Ok(None);
        };

        let factory = SlideFactory::new_shared_with_limits(
            &self.inner.powerpoint_document,
            &self.inner.persist_mapping,
            &self.inner.slide_directory,
            self.inner.source_limits,
        );
        let slide_data = factory.parse_slide(entry.persist_id())?;
        Ok(Some(
            Slide::from_slide_data(slide_data, position + 1)
                .with_shape_memo(&self.inner.shape_memo),
        ))
    }

    /// Get the number of slides (actual Slide records only).
    #[inline]
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.inner.slide_directory.len()
    }

    /// Return the validated logical slide directory.
    #[must_use]
    pub fn slide_directory(&self) -> &SlideDirectory {
        &self.inner.slide_directory
    }

    /// The theme of the first main master that carries a `PowerPoint` 2007
    /// `RoundTripThemeAtom` (record 0323); `None` for a presentation saved
    /// without one.
    ///
    /// # Errors
    ///
    /// Returns an error if a master's round-trip records or the theme part
    /// are malformed.
    pub fn theme(&self) -> Result<Option<crate::theme::Theme>> {
        for metadata in self.powerpoint12_main_master_metadata()? {
            if let Some(package) = &metadata.theme_package {
                return crate::theme::from_package(package).map(Some);
            }
        }
        Ok(None)
    }

    /// Parse `PowerPoint` 12 round-trip metadata for every main master slide.
    ///
    /// The returned values follow the main-master order in the `PowerPoint` document stream.
    /// Embedded ECMA-376 packages are validated but remain inert; no external resources or
    /// executable content are activated.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint12_main_master_metadata(&self) -> Result<Vec<MainMasterMetadata12>> {
        self.inner
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == RecordType::MainMaster)
            .map(MainMasterMetadata12::parse)
            .collect()
    }

    /// Return the typed slide-level programmable tags (MS-PPT 2.5.19) of every
    /// main master slide that carries a `SlideProgTagsContainer`, in
    /// main-master stream order.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ProgTags::slide_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn main_master_programmable_tags(&self) -> Result<Vec<crate::ProgTags>> {
        self.main_master_programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return main-master programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn main_master_programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Vec<crate::ProgTags>> {
        self.inner
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == RecordType::MainMaster)
            .filter_map(|record| crate::ProgTags::parse_slide(record, limits).transpose())
            .collect()
    }

    /// All header/footer metacharacter placeholders found in the document
    /// stream, in stream order.
    ///
    /// Slide-number, header, footer, and date placeholders (MS-PPT
    /// 2.9.47-2.9.52) mostly live in master text boxes. They are collected
    /// both from outline text bodies and from the Escher text boxes of slide,
    /// notes, and master drawing groups. They are inert: placeholders are
    /// never substituted, formatted, or laid out.
    ///
    /// Text-style defects repaired while the text boxes are parsed (see
    /// [`crate::TextPropDefect`]) are applied but *not* enumerable through
    /// this convenience method; use [`Self::text_metachars_with_report`] to
    /// enumerate them.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text_metachars(&self) -> Result<Vec<crate::text_metachar::TextMetachar>> {
        self.text_metachars_with_report(&mut crate::leniency::ToleranceReport::default())
    }

    /// All header/footer metacharacter placeholders found in the document
    /// stream, in stream order, recording every repaired non-structural
    /// text-style defect in `report`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text_metachars_with_report(
        &self,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Vec<crate::text_metachar::TextMetachar>> {
        use crate::text_metachar::TextMetachar;

        fn collect(
            record: &Record,
            out: &mut Vec<TextMetachar>,
            report: &mut crate::leniency::ToleranceReport,
            source_limits: SourceLimits,
        ) -> Result<()> {
            if record.record_type == RecordType::PPDrawing {
                for textbox in super::codec::drawing_textboxes(&record.data, source_limits)? {
                    let wrapper = crate::EscherTextboxWrapper::new_source_bound_with_report(
                        textbox.data().to_vec(),
                        source_limits,
                        report,
                    )?;
                    out.extend_from_slice(wrapper.metachars());
                }
            } else if matches!(
                record.record_type,
                RecordType::SlideNumberMCAtom
                    | RecordType::GenericDateMCAtom
                    | RecordType::HeaderMCAtom
                    | RecordType::FooterMCAtom
                    | RecordType::DateTimeMCAtom
                    | RecordType::RtfDateTimeMCAtom
            ) {
                out.extend(crate::text_metachar::metachars_from_records([record])?);
            }
            Ok(())
        }

        let mut result = Vec::new();
        for record in self.inner.parser.find_records_ref() {
            collect(record, &mut result, report, self.inner.source_limits)?;
        }
        Ok(result)
    }

    /// The normal three-pane view's splitter state (`NormalViewSetInfo9`,
    /// MS-PPT 2.4.21.2), when the document declares one. Files with multiple
    /// top-level Document containers yield the first occurrence.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn normal_view_set_info(&self) -> Result<Option<crate::view_set_info::NormalViewSet>> {
        for record in self.inner.parser.find_records_ref() {
            if record.record_type == RecordType::NormalViewSetInfo9 {
                return crate::view_set_info::NormalViewSet::parse_record(record).map(Some);
            }
        }
        Ok(None)
    }

    /// The notes-text view's scaling state (`NotesTextViewInfo9`, MS-PPT
    /// 2.4.21.4), when the document declares one. Files with multiple
    /// top-level Document containers yield the first occurrence.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn notes_text_view_info(&self) -> Result<Option<crate::view_set_info::NotesTextViewInfo>> {
        for record in self.inner.parser.find_records_ref() {
            if record.record_type == RecordType::NotesTextViewInfo9 {
                return crate::view_set_info::NotesTextViewInfo::parse_record(record).map(Some);
            }
        }
        Ok(None)
    }

    /// The default language and spelling settings for document text
    /// (`TextSIExceptionAtom`, MS-PPT 2.9.31), when declared.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text_special_info_defaults(
        &self,
    ) -> Result<Option<crate::text_si_exception::TextSpecialInfoDefaults>> {
        for record in self.inner.parser.find_records_ref() {
            if record.record_type == RecordType::TextSpecialInfoDefaultAtom {
                return crate::text_si_exception::TextSpecialInfoDefaults::parse_record(record)
                    .map(Some);
            }
        }
        Ok(None)
    }

    /// All outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78) in
    /// the document stream, in stream order.
    ///
    /// The atoms tie shape text boxes to outline text bodies and mostly live
    /// in the Escher text boxes of slide shapes, so they are collected both
    /// from the record tree and from every drawing-group text box.
    ///
    /// Text-style defects repaired while the text boxes are parsed (see
    /// [`crate::TextPropDefect`]) are applied but *not* enumerable through
    /// this convenience method; use [`Self::outline_text_refs_with_report`]
    /// to enumerate them.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn outline_text_refs(&self) -> Result<Vec<crate::OutlineTextRef>> {
        self.outline_text_refs_with_report(&mut crate::leniency::ToleranceReport::default())
    }

    /// All outline text references (`OutlineTextRefAtom`, MS-PPT 2.9.78) in
    /// the document stream, in stream order, recording every repaired
    /// non-structural text-style defect in `report`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn outline_text_refs_with_report(
        &self,
        report: &mut crate::leniency::ToleranceReport,
    ) -> Result<Vec<crate::OutlineTextRef>> {
        use crate::OutlineTextRef;

        fn collect(
            record: &Record,
            out: &mut Vec<OutlineTextRef>,
            report: &mut crate::leniency::ToleranceReport,
            source_limits: SourceLimits,
        ) -> Result<()> {
            if record.record_type == RecordType::PPDrawing {
                for textbox in super::codec::drawing_textboxes(&record.data, source_limits)? {
                    let wrapper = crate::EscherTextboxWrapper::new_source_bound_with_report(
                        textbox.data().to_vec(),
                        source_limits,
                        report,
                    )?;
                    out.extend_from_slice(wrapper.outline_text_refs());
                }
            } else if record.record_type == RecordType::OutlineTextRefAtom {
                out.push(OutlineTextRef::parse_record(record)?);
            }
            Ok(())
        }

        let mut result = Vec::new();
        for record in self.inner.parser.find_records_ref() {
            collect(record, &mut result, report, self.inner.source_limits)?;
        }
        Ok(result)
    }

    /// Parse document-level `PowerPoint` 12 settings and round-trip metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn powerpoint12_document_properties(&self) -> Result<DocumentProperties12> {
        let document = self.live_document_ref()?;
        DocumentProperties12::parse(document, &self.inner.source_limits)
    }

    /// Return the document routing-slip metadata, when present.
    ///
    /// Routing metadata is inert. This method never contacts recipients, starts
    /// a mail client, or updates routing status.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn routing_slip(&self) -> Result<Option<Slip>> {
        let document = self.live_document_ref()?;
        let mut slip_records = document
            .children
            .iter()
            .filter(|record| record.record_type == RecordType::DocRoutingSlipAtom);
        let Some(record) = slip_records.next() else {
            return Ok(None);
        };
        if slip_records.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple routing slips".to_string(),
            ));
        }
        Slip::parse(record).map(Some)
    }

    /// Return the strictly validated document-level embedded sound collection.
    ///
    /// Sound bytes remain borrowed and inert. This method never decodes audio,
    /// invokes a codec, plays media, or resolves an external resource.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn embedded_sounds(&self) -> Result<Option<Collection<'_>>> {
        let document = self.live_document_ref()?;
        let mut collections = document
            .children
            .iter()
            .filter(|record| record.record_type == RecordType::SoundCollection);
        let Some(collection) = collections.next() else {
            return Ok(None);
        };
        if collections.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple SoundCollection containers".to_string(),
            ));
        }
        Collection::parse(collection).map(Some)
    }

    /// Validate every shape, text-range, and outline-text sound reference.
    ///
    /// This performs no audio decoding or playback. Null references are valid;
    /// every non-null reference must resolve in the document `SoundCollection`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn validate_interaction_sound_references(&self) -> Result<()> {
        let sounds = self.embedded_sounds()?;
        let validate = |interaction: &crate::Interaction| -> Result<()> {
            if interaction.sound_id == 0 {
                return Ok(());
            }
            let sound_collection = sounds.as_ref().ok_or_else(|| {
                Error::Corrupted(
                    "interaction references a sound but the document has no SoundCollection"
                        .to_string(),
                )
            })?;
            interaction.validate_sound_collection(sound_collection)
        };

        for slide in self.slides()? {
            for entry in slide.shape_interactions()? {
                for interaction in &entry.interactions {
                    validate(interaction)?;
                }
            }
            for entry in slide.shape_text_interactions()? {
                for interaction in &entry.interactions {
                    validate(&interaction.interaction)?;
                }
            }
            for body in slide.outline_text_interactions() {
                for interaction in &body.interactions {
                    validate(&interaction.interaction)?;
                }
            }
        }
        Ok(())
    }

    /// Return strictly validated inert audio/video metadata from `ExObjListContainer`.
    ///
    /// Paths are never accessed and embedded sound bytes are never decoded or played.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn external_media(&self) -> Result<Option<MediaCollection>> {
        let document = self.live_document_ref()?;
        let source_limits = self.inner.source_limits;
        let media_limits = MediaLimits {
            max_root_bytes: source_limits
                .max_record_bytes()
                .min(source_limits.max_input_bytes()),
            max_record_bytes: source_limits.max_record_bytes(),
            max_depth: source_limits.max_depth(),
            max_records: source_limits.max_events(),
            max_owner_references: source_limits.max_events(),
        };
        let Some(media) = MediaCollection::parse(document, &media_limits)? else {
            return Ok(None);
        };
        let _ = OleCollection::parse(document, &source_limits)?;
        let sounds = self.embedded_sounds()?;
        media.validate_sound_collection(sounds.as_ref())?;
        Ok(Some(media))
    }

    /// Return the document's strictly validated hyperlink definitions.
    ///
    /// Targets remain inert: this method never opens a URL, path, presentation,
    /// or named show.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn hyperlinks(&self) -> Result<Hyperlinks> {
        let source_limits = self.inner.source_limits;
        Hyperlinks::parse(&self.live_document_record()?, &source_limits)
    }

    /// Resolve a click interaction (record 0323): a URL, a slide of this
    /// presentation or a slide-show jump, with its screen tip. `None` when
    /// the action is not a hyperlink.
    ///
    /// # Errors
    ///
    /// Returns an error if the document's hyperlink records are malformed.
    pub fn resolve_click(
        &self,
        interaction: &crate::hyperlink::Interaction,
    ) -> Result<Option<crate::hyperlink::Click>> {
        let hyperlinks = self.hyperlinks()?;
        let slide_ids = self
            .slide_directory()
            .entries()
            .iter()
            .map(crate::SlideDirectoryEntry::slide_id)
            .collect::<Vec<_>>();
        Ok(crate::hyperlink::Click::resolve(
            interaction,
            &hyperlinks,
            &slide_ids,
        ))
    }

    /// The click hyperlink of the shape with `OfficeArt` ID `shape_id` on
    /// the slide at zero-based `position`; see [`Self::resolve_click`].
    ///
    /// # Errors
    ///
    /// Returns an error if the slide or the document's hyperlink records are
    /// malformed.
    pub fn shape_click(
        &self,
        position: usize,
        shape_id: u32,
    ) -> Result<Option<crate::hyperlink::Click>> {
        let Some(slide) = self.slide_at(position)? else {
            return Ok(None);
        };
        let Some(entry) = slide
            .shape_interactions()?
            .into_iter()
            .find(|entry| entry.shape_id == shape_id)
        else {
            return Ok(None);
        };
        let hyperlinks = self.hyperlinks()?;
        let slide_ids = self
            .slide_directory()
            .entries()
            .iter()
            .map(crate::SlideDirectoryEntry::slide_id)
            .collect::<Vec<_>>();
        Ok(entry.interactions.iter().find_map(|interaction| {
            crate::hyperlink::Click::resolve(interaction, &hyperlinks, &slide_ids)
        }))
    }

    /// Return inert embedded and linked OLE metadata without loading object storage.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn ole_objects(&self) -> Result<Option<OleCollection>> {
        let document = self.live_document_ref()?;
        let source_limits = self.inner.source_limits;
        let Some(objects) = OleCollection::parse(document, &source_limits)? else {
            return Ok(None);
        };
        objects.validate_persist_mapping(&self.inner.persist_mapping)?;
        Ok(Some(objects))
    }

    /// Resolve a persisted embedded OLE-object storage as bounded inert bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn ole_storage(&self, persist_id: u32) -> Result<Option<Storage>> {
        self.ole_storage_as(persist_id, StorageKind::OleObject)
    }

    /// Resolve a persisted storage with the kind supplied by its referencing record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn ole_storage_as(&self, persist_id: u32, kind: StorageKind) -> Result<Option<Storage>> {
        let Some(offset) = self.inner.persist_mapping.get_offset(persist_id) else {
            return Ok(None);
        };
        let storage_offset = usize::try_from(offset)
            .map_err(|_err| Error::Corrupted("OLE storage offset exceeds usize".to_string()))?;
        let (record, _) = Record::parse(
            &self.inner.powerpoint_document,
            storage_offset,
            &self.inner.source_limits,
        )?;
        if record.record_type != RecordType::ExternalOleObjectStg {
            return Err(Error::Corrupted(format!(
                "persist ID {persist_id} does not reference ExOleObjStg"
            )));
        }
        Storage::parse_as(&record, kind).map(Some)
    }

    /// Enumerate inert native charts as neutral Office Graph views.
    ///
    /// Embedded objects whose subtype or `ProgID` identifies `MSGraph.Chart` or
    /// `Excel.Chart` ([MS-PPT] 2.13.11) have their `ExOleObjStg` payload opened
    /// as a bounded compound storage and their chart substreams validated by
    /// `litchi-ograph`. Linked charts are never opened. A corrupt payload
    /// degrades to a per-object failure without aborting the remaining charts.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn charts(&self) -> Result<crate::chart::Inventory> {
        self.charts_with(litchi_ograph::Limits::default())
    }

    /// Enumerate native charts with explicit neutral resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn charts_with(&self, limits: litchi_ograph::Limits) -> Result<crate::chart::Inventory> {
        crate::chart::enumerate(self, limits)
    }

    /// Return the live document-comparison snapshot.
    ///
    /// Review metadata is inert: this accessor never compares presentations,
    /// opens the embedded reviewer document, or applies/rejects a change.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn document_comparison(&self) -> Result<crate::document_comparison::Snapshot> {
        crate::document_comparison::package::from_presentation(self)
    }

    /// Parse the base and `PowerPoint` 10 font owners from the exact live
    /// `DocumentContainer`. Embedded EOT payloads remain inert.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn fonts(&self) -> Result<crate::font::FontCollections> {
        self.fonts_with_limits(crate::font::Limits::default())
    }

    /// Parse live font owners with explicit composed record/font limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn fonts_with_limits(
        &self,
        mut limits: crate::font::Limits,
    ) -> Result<crate::font::FontCollections> {
        limits.records = limits.records.constrained_by(self.inner.source_limits);
        crate::font::FontCollections::parse_with_limits(&self.live_document_record()?, limits)
    }

    /// Resolve the live `DocumentContainer` record via the persist directory.
    ///
    /// Incrementally saved presentations can hold several `DocumentContainer`
    /// records; only the one referenced by the current `UserEditAtom` is live.
    /// The parsed live `DocumentContainer`: the one the current persist
    /// directory maps. An incrementally saved presentation keeps superseded
    /// `DocumentContainer` records in the stream (record 0357).
    pub(crate) fn live_document_ref(&self) -> Result<&Record> {
        let live = self.live_document_record()?;
        self.inner
            .parser
            .find_records_ref()
            .into_iter()
            .find(|record| record.record_type == RecordType::Document && **record == live)
            .ok_or_else(|| {
                Error::Corrupted("PowerPoint live Document container was not parsed".to_string())
            })
    }

    pub(crate) fn live_document_record(&self) -> Result<Record> {
        self.check_source()?;
        let persist_id = self.inner.slide_directory.document_persist_id();
        let offset = self
            .inner
            .persist_mapping
            .get_offset(persist_id)
            .ok_or_else(|| {
                Error::Corrupted(format!("document persist ID {persist_id} has no mapping"))
            })?;
        let document_offset = usize::try_from(offset)
            .map_err(|_err| Error::Corrupted("document offset exceeds usize".to_string()))?;
        let (record, _) = Record::parse(
            &self.inner.powerpoint_document,
            document_offset,
            &self.inner.source_limits,
        )?;
        if record.record_type != RecordType::Document {
            return Err(Error::Corrupted(
                "document persist ID does not resolve to a DocumentContainer".to_string(),
            ));
        }
        Ok(record)
    }

    /// The exact encoded bytes of the live `DocumentContainer` record.
    pub(crate) fn live_document_bytes(&self) -> Result<Vec<u8>> {
        self.check_source()?;
        let offset = self.inner.slide_directory.document_offset();
        let (record, consumed) = Record::parse(
            &self.inner.powerpoint_document,
            offset,
            &self.inner.source_limits,
        )?;
        if record.record_type != RecordType::Document {
            return Err(Error::Corrupted(
                "document persist ID does not resolve to a DocumentContainer".to_string(),
            ));
        }
        offset
            .checked_add(consumed)
            .and_then(|end| self.inner.powerpoint_document.get(offset..end))
            .map(<[u8]>::to_vec)
            .ok_or_else(|| Error::Corrupted("live DocumentContainer exceeds its stream".into()))
    }

    pub(crate) fn document_stream(&self) -> &[u8] {
        &self.inner.powerpoint_document
    }

    /// Return a payload-free descriptor for the document's VBA project storage.
    ///
    /// Macro bytes are not returned, decompressed, interpreted, or executed.
    /// A non-null persist reference must resolve to a `VbaProjectStg` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn vba_project_storage(&self) -> Result<Option<crate::VbaProjectStorage>> {
        let records = self.inner.parser.find_records_ref();
        let Some(info) = crate::VbaInfo::parse_records(&records)? else {
            return Ok(None);
        };
        let storage = self.resolve_vba_storage(info)?;
        crate::VbaProjectStorage::from_info_and_metadata(info, storage.map(StorageRef::metadata))
            .map(Some)
    }

    /// Return validated inert `VBAInfoAtom` metadata for the document.
    ///
    /// For richer outer-storage metadata without exposing VBA payload bytes,
    /// use [`Self::vba_project_storage`].
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn vba_info(&self) -> Result<Option<crate::VbaInfo>> {
        Ok(self.vba_project_storage()?.map(|storage| storage.info()))
    }

    /// Parse the embedded MS-OVBA project with the limits retained by the
    /// owning open options.
    ///
    /// The outer zlib stream, inner CFB, and MS-OVBA compressed containers are
    /// bounded, and source is never compiled, interpreted, or executed. Set
    /// [`crate::OpenOptions::vba_project_limits`] before constructing this
    /// immutable projection to select custom ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(feature = "vba-inspection")]
    pub fn vba(
        &self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, crate::VbaProjectError> {
        let limits = self.inner.vba_project_limits;
        let records = self.inner.parser.find_records_ref();
        let Some(info) = crate::VbaInfo::parse_records(&records)? else {
            return Ok(None);
        };
        let maybe_storage = self.resolve_vba_storage(info)?;
        let summary = crate::VbaProjectStorage::from_info_and_metadata(
            info,
            maybe_storage.map(StorageRef::metadata),
        )?;
        if !summary.has_persisted_storage() {
            return Ok(None);
        }
        let Some(storage_ref) = maybe_storage else {
            return Err(Error::Corrupted(format!(
                "VBAInfoAtom persist ID {} has no storage record",
                summary.persist_id_ref()
            ))
            .into());
        };
        let cfb_limits = crate::vba_info::vba_cfb_limits(&limits)?;
        let storage = storage_ref.check_stored_limit(limits.max_stored_bytes)?;
        let decompressed_limit = limits.max_cfb_bytes.min(limits.project.max_cfb_bytes);
        let cfb = storage.decompressed_bytes(decompressed_limit)?;
        let mut ole = OleFile::open(Cursor::new(cfb.as_ref()), cfb_limits)
            .map_err(litchi_vba::Error::from)
            .map_err(crate::VbaProjectError::from)?;
        litchi_vba::project::Project::open(&mut ole, &[], &limits.project)
            .map(Some)
            .map_err(Into::into)
    }

    #[cfg(feature = "vba-inspection")]
    fn resolve_vba_storage(&self, info: crate::VbaInfo) -> Result<Option<StorageRef<'_>>> {
        if info.persist_id_ref == 0 {
            return Ok(None);
        }
        let offset = self
            .inner
            .persist_mapping
            .get_offset(info.persist_id_ref)
            .ok_or_else(|| {
                Error::Corrupted(format!(
                    "VBAInfoAtom persist ID {} has no storage record",
                    info.persist_id_ref
                ))
            })?;
        let storage_offset = usize::try_from(offset)
            .map_err(|_err| Error::Corrupted("VBA storage offset exceeds usize".to_string()))?;
        StorageRef::parse_at(
            &self.inner.powerpoint_document,
            storage_offset,
            StorageKind::VbaProject,
        )
        .map(Some)
    }

    /// Return the PPT10 modify-password metadata without verifying it.
    ///
    /// The returned value redacts its secret from `Debug`; callers must use an
    /// explicitly named accessor to inspect it. This method does not decrypt
    /// the presentation or grant modification access.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn modify_password(&self) -> Result<Option<crate::ModifyPassword>> {
        let document = self.live_document_ref()?;
        crate::ModifyPassword::parse_document(document, &self.inner.source_limits)
    }

    /// Return the inert `PowerPoint` 10 privacy preference, if present.
    ///
    /// This accessor never removes document metadata or rewrites the file.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn privacy_settings(&self) -> Result<Option<crate::PrivacySettings>> {
        let document = self.live_document_ref()?;
        crate::PrivacySettings::parse_document(document, &self.inner.source_limits)
    }

    /// Return typed `PowerPoint` 9 Presentation Advisor warning preferences.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn presentation_advisor_settings(
        &self,
    ) -> Result<Option<crate::PresentationAdvisorSettings>> {
        let document = self.live_document_ref()?;
        crate::PresentationAdvisorSettings::parse_document(document, &self.inner.source_limits)
    }

    /// Return typed `PowerPoint` 9 Web-document publishing preferences.
    ///
    /// This accessor never writes files, invokes a browser, or exports content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn html_document_settings(&self) -> Result<Option<crate::HtmlDocumentSettings>> {
        let document = self.live_document_ref()?;
        crate::HtmlDocumentSettings::parse_document(document, &self.inner.source_limits)
    }

    /// Return inert `PowerPoint` 9 Web-publication metadata, if present.
    ///
    /// Named-show references are cross-validated against the presentation's
    /// named-show table. This accessor never writes files, resolves a URI,
    /// invokes a browser, or exports presentation content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn html_publish_settings(&self) -> Result<Option<crate::HtmlPublishSettings>> {
        let document = self.live_document_ref()?;
        crate::HtmlPublishSettings::parse_document(document, &self.inner.source_limits)
    }

    /// Return all inert `PowerPoint` 9 presentation-broadcast descriptions.
    ///
    /// This accessor never contacts a server, opens a URL or ASD file, sends
    /// mail, records media, or starts a broadcast.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn broadcasts(&self) -> Result<crate::Broadcasts> {
        // The live document: an incrementally saved presentation keeps its
        // superseded `DocumentContainer` records in the stream (record 0357).
        crate::Broadcasts::parse_document(&self.live_document_record()?, &self.inner.source_limits)
    }

    /// Return the structurally decoded, inert `PowerPoint` 9 mail envelope.
    ///
    /// This accessor never sends mail, invokes a mail client, opens an
    /// attachment, or evaluates attachment bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn envelope_data(&self) -> Result<Option<crate::EnvelopeData>> {
        let document = self.live_document_ref()?;
        crate::EnvelopeData::parse_document(document, &self.inner.source_limits)
    }

    /// Validate and expose the specification-defined terminal document records.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn document_structure(&self) -> Result<crate::DocumentStructure> {
        let document = self.live_document_ref()?;
        crate::DocumentStructure::parse(document)
    }

    /// Return the document-wide `DocumentAtom` (MS-PPT 2.4.2), if present.
    ///
    /// Slide geometry, OLE server zoom, master persist references, and display
    /// flags are inert metadata and nothing is rendered. Incremental histories
    /// are resolved through the live document persist mapping.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn document_atom(&self) -> Result<Option<crate::DocumentAtom>> {
        let document = self.live_document_record()?;
        let mut atoms = document
            .children
            .iter()
            .filter(|record| record.record_type == RecordType::DocumentAtom);
        let Some(atom) = atoms.next() else {
            return Ok(None);
        };
        if atoms.next().is_some() {
            return Err(Error::Corrupted(
                "live DocumentContainer contains multiple DocumentAtom records".into(),
            ));
        }
        crate::DocumentAtom::parse(atom).map(Some)
    }

    /// Collect the color-scheme atoms of every slide, notes, main master, and
    /// handout container in stream order (MS-PPT 2.5.14, 2.5.15).
    ///
    /// The colors are inert display metadata: nothing is rendered, resolved
    /// against a theme, or applied to shapes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn color_schemes(&self) -> Result<Vec<crate::ColorSchemeAtom>> {
        let mut schemes = Vec::new();
        for record in self.inner.parser.find_records_ref() {
            if matches!(
                record.record_type,
                RecordType::Slide
                    | RecordType::Notes
                    | RecordType::MainMaster
                    | RecordType::Handout
            ) {
                schemes.extend(crate::ColorSchemeAtom::collect(record)?);
            }
        }
        Ok(schemes)
    }

    /// Return inert `PowerPoint` 9 mail-envelope state, if present.
    ///
    /// This accessor never sends mail, invokes a mail client, or interprets the
    /// associated envelope payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn envelope_settings(&self) -> Result<Option<crate::EnvelopeSettings>> {
        let document = self.live_document_ref()?;
        crate::EnvelopeSettings::parse_document(document, &self.inner.source_limits)
    }

    /// Return strictly validated header/footer metadata from all specification-defined scopes.
    ///
    /// Date identifiers and text remain inert metadata. This method does not format dates,
    /// execute content, resolve resources, or modify the underlying OLE file.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn header_footers(&self) -> Result<HeaderFooters> {
        let records = self.inner.parser.find_records_ref();
        let mut values = HeaderFooters::parse_record_tree(&records)?;

        let mut first_master_display = None;
        for (master_ordinal, master) in records
            .iter()
            .filter(|record| record.record_type == RecordType::MainMaster)
            .enumerate()
        {
            if let Some(display) =
                super::codec::placeholder_display_from_record(master, self.inner.source_limits)?
            {
                let scope = HeaderFooterScope::Local {
                    parent: HeaderFooterParent::MainMaster,
                    parent_ordinal: HeaderFooterParentOrdinal::new(master_ordinal),
                };
                if values.has_scope(scope) {
                    values.attach_placeholder_display(scope, display.clone())?;
                }
                if first_master_display.is_none() {
                    first_master_display = Some(display);
                }
            }
        }
        let has_master_display = first_master_display.is_some();
        if let Some(display) = first_master_display {
            values.attach_placeholder_display(HeaderFooterScope::PresentationSlides, display)?;
        }

        let mut first_unoverridden_slide_display = None;
        let mut first_notes_display = None;
        for (ordinal, slide) in records
            .iter()
            .filter(|record| record.record_type == RecordType::Slide)
            .enumerate()
        {
            let scope = HeaderFooterScope::Local {
                parent: HeaderFooterParent::Slide,
                parent_ordinal: HeaderFooterParentOrdinal::new(ordinal),
            };
            if let Some(display) =
                super::codec::placeholder_display_from_record(slide, self.inner.source_limits)?
            {
                if values.has_scope(scope) {
                    values.attach_placeholder_display(scope, display)?;
                } else {
                    if first_unoverridden_slide_display.is_none() {
                        first_unoverridden_slide_display = Some(display.clone());
                    }
                    values.attach_placeholder_display(scope, display)?;
                }
            }
        }
        for slide in self.slides()? {
            if first_notes_display.is_some() {
                break;
            }
            if let Some(notes) = slide.speaker_notes()? {
                first_notes_display =
                    super::codec::placeholder_display_from_shapes(notes.shapes()?)?;
            }
        }
        if !has_master_display && let Some(display) = first_unoverridden_slide_display {
            values.attach_placeholder_display(HeaderFooterScope::PresentationSlides, display)?;
        }
        if let Some(display) = first_notes_display {
            values.attach_placeholder_display(HeaderFooterScope::NotesAndHandouts, display)?;
        }
        Ok(values)
    }

    /// Extract all text from the presentation.
    ///
    /// # Performance
    ///
    /// - Iterates through all slides
    /// - Each slide extracts text lazily
    /// - Text is collected and joined
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn text(&self) -> Result<String> {
        self.check_source()?;
        let factory = SlideFactory::new_shared_with_limits(
            &self.inner.powerpoint_document,
            &self.inner.persist_mapping,
            &self.inner.slide_directory,
            self.inner.source_limits,
        );
        let mut text = String::new();

        for (idx, entry) in self.inner.slide_directory.entries().iter().enumerate() {
            let slide = Slide::from_slide_data(factory.parse_slide(entry.persist_id())?, idx + 1);
            let Ok(slide_text) = slide.text() else {
                continue;
            };
            if slide_text.is_empty() {
                continue;
            }
            if !text.is_empty() {
                text.push_str("\n\n");
            }
            text.push_str(slide_text);
        }

        Ok(if text.is_empty() {
            String::from("No text content found in presentation")
        } else {
            text
        })
    }

    /// Fast text extraction that skips shape parsing.
    ///
    /// This is optimized for cases where only text is needed (e.g., markdown conversion)
    /// and shape information is not required.
    ///
    /// # Performance
    ///
    /// - Directly extracts text from slide records without parsing shapes
    /// - Significantly faster than `slides()` + `text()` for large presentations
    /// - No shape object allocation or geometry calculations
    /// - Pre-allocated string buffer
    ///
    /// # Returns
    ///
    /// Vector of (`slide_number`, text) tuples for each slide
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn extract_text_fast(&self) -> Result<Vec<(usize, String)>> {
        self.check_source()?;
        let factory = SlideFactory::new_shared_with_limits(
            &self.inner.powerpoint_document,
            &self.inner.persist_mapping,
            &self.inner.slide_directory,
            self.inner.source_limits,
        );

        let mut results = Vec::with_capacity(self.inner.slide_directory.len());

        for (idx, slide_result) in factory.slides().enumerate() {
            let slide_data = slide_result?;

            // Pre-allocate string buffer
            let mut text = String::with_capacity(512);

            // Extract text from slide records without parsing shapes
            let record_text = slide_data.record.extract_text()?;
            let trimmed = record_text.trim();
            if !trimmed.is_empty() {
                text.push_str(trimmed);
            }

            // Extract text from Escher/PPDrawing using the optimized path,
            // resolving outline references as `Slide::text` does (L-PPT-R1)
            if let Some(ppdrawing) = slide_data.record.find_child(RecordType::PPDrawing) {
                let escher_text = crate::odraw::text_from_slide_drawing(
                    &ppdrawing.data,
                    slide_data.outline_text_bodies,
                    self.inner.source_limits,
                )?;
                let escher_trimmed = escher_text.trim();
                if !escher_trimmed.is_empty() {
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(escher_trimmed);
                }
            }

            if text.is_empty() {
                text.push_str(&slide_data.slide_list_text);
            }

            results.push((idx + 1, text));
        }

        Ok(results)
    }

    /// Extract all images from the presentation
    ///
    /// Images are returned in semantic `BStore` order. FBSE metadata comes from
    /// the document's `PPDrawingGroup`; delayed payloads are resolved against
    /// the headerless `Pictures` stream.
    ///
    /// # Returns
    /// Vector of all extracted images with metadata
    ///
    /// # Example
    /// ```no_run
    /// use litchi_ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for image in pres.images()? {
    ///     std::fs::write(image.filename(), image.data()?)?;
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn images(&self) -> Result<Vec<ImageFile<'_>>> {
        self.check_source()?;
        if let Some(store) = self.image_store()? {
            return litchi_odraw::image::all(&store, self.pictures_bytes()?).map_err(Error::from);
        }
        self.pictures_bytes()?.map_or_else(
            || Ok(Vec::new()),
            |pictures| {
                litchi_odraw::image::delay(pictures, litchi_odraw::image::Limits::default())
                    .map_err(Error::from)
            },
        )
    }

    /// Extract an image by BLIP ID
    ///
    /// This method is used internally by `PictureShape` to resolve
    /// BLIP ID references to actual image data.
    ///
    /// # Arguments
    /// * `id` - Checked BLIP ID from the shape's Escher properties
    ///
    /// # Returns
    /// The extracted image, or None if not found
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn image(&self, id: ImageId) -> Result<Option<ImageFile<'_>>> {
        self.check_source()?;
        let Some(store) = self.image_store()? else {
            return Ok(None);
        };
        litchi_odraw::image::get(&store, id, self.pictures_bytes()?).map_err(Error::from)
    }

    /// Resolves a raw one-based host index after checking its `OfficeArt` range.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn image_at(&self, index: u32) -> Result<Option<ImageFile<'_>>> {
        self.check_source()?;
        self.image(ImageId::new(index)?)
    }

    /// Parses the borrowed BLIP store from the document's drawing group.
    ///
    /// No parsed store is retained inside `Presentation`; this avoids
    /// self-referential storage while keeping the view zero-copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn image_store(&self) -> Result<Option<ImageStore<'_>>> {
        self.check_source()?;
        let mut store = None;
        for record in self
            .inner
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == RecordType::PPDrawingGroup)
        {
            let candidate =
                litchi_odraw::image::store(&record.data, litchi_odraw::image::Limits::default())
                    .map_err(Error::from)?;
            if let Some(parsed_store) = candidate
                && store.replace(parsed_store).is_some()
            {
                return Err(Error::Corrupted(
                    "Presentation contains multiple OfficeArt BStore containers".to_string(),
                ));
            }
        }
        Ok(store)
    }

    fn pictures_bytes(&self) -> Result<Option<&[u8]>> {
        if let Some(data) = self.inner.pictures_data.as_deref() {
            return Ok(Some(data));
        }
        self.inner
            .pictures_source
            .as_ref()
            .map_or(Ok(None), |source| source.bytes().map(Some))
    }

    /// Check if the presentation has a Pictures stream
    #[must_use]
    pub fn has_pictures(&self) -> bool {
        self.inner.pictures_data.is_some() || self.inner.pictures_source.is_some()
    }

    /// Strictly parse slide and notes editing-view information.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn slide_view_information(&self) -> Result<SlideViewInformation> {
        let records = self.inner.parser.find_records_ref();
        SlideViewInformation::parse_records(&records)
    }

    /// Strictly parse outline and slide-sorter editing-view information.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn outline_sorter_view_information(&self) -> Result<OutlineSorterViewInformation> {
        let records = self.inner.parser.find_records_ref();
        OutlineSorterViewInformation::parse_records(&records)
    }

    /// Parse all slide comments in the presentation.
    ///
    /// Comments are stored per slide inside `ProgTags/ProgBinaryTag/BinaryTagData`
    /// as `Comment2000` (type=12000) containers. Only slides with at least one
    /// comment produce an entry.
    ///
    /// # Errors
    ///
    /// Returns an error when a slide or one of its comment records is malformed.
    pub fn comments(&self) -> Result<Vec<ParsedSlideComments>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            let comments = slide.comments()?;
            if !comments.is_empty() {
                result.push(ParsedSlideComments {
                    slide_number: slide.slide_number(),
                    comments,
                });
            }
        }
        Ok(result)
    }

    /// Parse and validate the complete inert presentation-comment catalog.
    ///
    /// Author seed metadata lives in the document stream while comment atoms
    /// live in slide extensions; this facade joins both scopes and checks the
    /// cross-record index rule before returning the inventory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn comment_catalog(&self) -> Result<crate::comments::Catalog> {
        let authors = crate::comments::Authors::parse(
            &self.live_document_record()?,
            &self.inner.source_limits,
        )?;
        crate::comments::Catalog::from_parts(authors, self.comments()?)
    }

    /// Return all shape-scoped programmable tags in slide and shape order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_programmable_tags(
        &self,
    ) -> Result<Vec<crate::PresentationShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    /// Return the inert document-wide `PowerPoint` 11 smart-tag store.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn smart_tags(&self) -> Result<Option<crate::SmartTagStore>> {
        let document = self.live_document_record()?;
        crate::SmartTagStore::parse(&document, &self.inner.source_limits)
    }

    /// Return the typed document-level programmable tags (MS-PPT 2.4.23), when
    /// the document carries a `DocProgTagsContainer`.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ProgTags::document_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn programmable_tags(&self) -> Result<Option<crate::ProgTags>> {
        self.programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return document-level programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Option<crate::ProgTags>> {
        crate::ProgTags::parse_document(&self.live_document_record()?, limits)
    }

    /// Return all shape programmable tags with caller-supplied resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_programmable_tags_with_limits(
        &self,
        limits: crate::ShapeProgrammableTagLimits,
    ) -> Result<Vec<crate::PresentationShapeProgrammableTagsEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.shape_programmable_tags_with_limits(limits)? {
                result.push(crate::PresentationShapeProgrammableTagsEntry {
                    slide_number: slide.slide_number(),
                    shape_id: entry.shape_id,
                    programmable_tags: entry.programmable_tags,
                });
            }
        }
        Ok(result)
    }

    /// Return every typed shape-flag projection in slide and shape order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn shape_flags(&self) -> Result<Vec<crate::PresentationShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn shape_flags_with_limits(
        &self,
        limits: crate::ShapeFlagLimits,
    ) -> Result<Vec<crate::PresentationShapeFlagEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.shape_flags_with_limits(limits)? {
                result.push(crate::PresentationShapeFlagEntry {
                    slide_number: slide.slide_number(),
                    shape_id: entry.shape_id,
                    projection: entry.projection,
                });
            }
        }
        Ok(result)
    }

    /// Return every context-validated presentation-slide placeholder.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn placeholder_atoms(&self) -> Result<Vec<crate::PresentationPlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::PlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn placeholder_atoms_with_limits(
        &self,
        limits: crate::PlaceholderLimits,
    ) -> Result<Vec<crate::PresentationPlaceholderEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.placeholder_atoms_with_limits(limits)? {
                result.push(crate::PresentationPlaceholderEntry {
                    slide_number: slide.slide_number(),
                    shape_id: entry.shape_id,
                    placeholder: entry.placeholder,
                });
            }
        }
        Ok(result)
    }

    /// Parse custom slide shows (named shows) from the Document container.
    ///
    /// Custom shows are stored as `NamedShows` (type=1040) container in the
    /// Document record, containing `NamedShow` (type=1041) children with
    /// `CString` names and `NamedShowSlides` (type=1042) slide ID arrays.
    ///
    /// # Returns
    ///
    /// A vector of `(name, slide_indices)` tuples for each custom show.
    /// Slide indices are 0-based.
    #[must_use]
    pub fn custom_shows(&self) -> Vec<ParsedCustomShow> {
        let mut shows = Vec::new();

        // Parse Document record from the stream
        let records = self.inner.parser.find_records_ref();
        for record in &records {
            if record.record_type == RecordType::Document {
                // Find NamedShows container in Document
                for child in &record.children {
                    if child.record_type == RecordType::NamedShows {
                        Self::parse_named_shows(child, &mut shows);
                    }
                }
            }
        }

        shows
    }

    /// Parse `NamedShow` containers from a `NamedShows` container.
    pub(super) fn parse_named_shows(named_shows: &Record, shows: &mut Vec<ParsedCustomShow>) {
        for child in &named_shows.children {
            if child.record_type == RecordType::NamedShow {
                let mut name = String::new();
                let mut slide_indices = Vec::new();

                for sub in &child.children {
                    if sub.record_type == RecordType::CString {
                        // UTF-16LE name
                        let chars: Vec<u16> = sub
                            .data
                            .as_chunks::<2>()
                            .0
                            .iter()
                            .map(|c| u16::from_le_bytes([c[0], c[1]]))
                            .collect();
                        name = String::from_utf16_lossy(&chars);
                    } else if sub.record_type == RecordType::NamedShowSlides {
                        // Array of u32 slide IDs (0x100 + slide_index)
                        for chunk in sub.data.as_chunks::<4>().0.iter() {
                            let slide_id =
                                u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
                            // Convert slide ID (0x100+index) back to 0-based index
                            let index = slide_id.saturating_sub(0x100) as usize;
                            slide_indices.push(index);
                        }
                    }
                }

                if !name.is_empty() {
                    shows.push(ParsedCustomShow {
                        name,
                        slide_indices,
                    });
                }
            }
        }
    }
}

/// A parsed custom slide show from a PPT file.
#[derive(Debug, Clone)]
pub struct ParsedCustomShow {
    /// Show name.
    pub name: String,
    /// 0-based slide indices in presentation order.
    pub slide_indices: Vec<usize>,
}

/// Comments parsed from a single slide of a presentation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParsedSlideComments {
    /// 1-based slide number the comments belong to.
    pub slide_number: usize,
    /// Comments in record order.
    pub comments: Vec<ParsedComment>,
}
