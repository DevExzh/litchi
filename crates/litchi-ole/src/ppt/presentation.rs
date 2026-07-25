use super::super::OleFile;
use super::document_properties::PowerPoint12DocumentProperties;
#[cfg(feature = "imgconv")]
use super::encryption::decrypt_pictures;
use super::encryption::decrypt_powerpoint_document;
use super::external_media::PowerPointExternalMediaCollection;
use super::header_footer::{
    PowerPointHeaderFooterDisplayText, PowerPointHeaderFooterParent,
    PowerPointHeaderFooterParentOrdinal, PowerPointHeaderFooterScope, PowerPointHeaderFooters,
};
use super::main_master::PowerPoint12MainMasterMetadata;
use super::non_zoom_view::PowerPointOutlineSorterViewInformation;
use super::ole_object::PowerPointOleObjectCollection;
use super::ole_storage::PowerPointOleStorage;
/// High-performance Presentation API with zero-copy slide parsing.
use super::package::{PptError, PptOpenOptions, Result};
use super::parsers::PptRecordParser;
use super::persist::PersistMapping;
use super::records::PptRecord;
use super::routing_slip::PowerPointRoutingSlip;
use super::slide::{ParsedComment, Slide, SlideDirectory, SlideFactory};
use super::sound_collection::PowerPointSoundCollection;
use super::view_info::PowerPointSlideViewInformation;
use crate::consts::PptRecordType;
#[cfg(feature = "imgconv")]
use crate::extractor::{ExtractedImage, ImageExtractor};
#[cfg(feature = "imgconv")]
use litchi_imgconv::BlipStore;
use std::io::{Read, Seek};

/// A PowerPoint presentation (.ppt) with high-performance zero-copy parsing.
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
/// use litchi_ole::ppt::Package;
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
pub struct Presentation {
    /// The main document stream data (owned for lifetime management)
    powerpoint_document: Vec<u8>,
    /// Parsed record structure (reserved for future advanced parsing)
    #[allow(dead_code)]
    pub(crate) parser: PptRecordParser,
    /// Persist ID to offset mapping
    pub(crate) persist_mapping: PersistMapping,
    slide_directory: SlideDirectory,
    /// Pictures stream data (for image extraction)
    #[cfg(feature = "imgconv")]
    pictures_data: Option<Vec<u8>>,
    /// BLIP store (image metadata index)
    #[cfg(feature = "imgconv")]
    blip_store: Option<BlipStore<'static>>,
}

impl Presentation {
    /// Create a new Presentation from an OLE file.
    pub(crate) fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self> {
        Self::from_ole_with_options(ole, PptOpenOptions::default())
    }

    /// Create a new Presentation from an OLE file with password-to-open options.
    pub(crate) fn from_ole_with_options<R: Read + Seek>(
        ole: &mut OleFile<R>,
        options: PptOpenOptions<'_>,
    ) -> Result<Self> {
        // Read the PowerPoint Document stream
        let mut powerpoint_document = Self::read_powerpoint_document(ole)?;
        let current_user_data = ole
            .open_stream(&["Current User"])
            .or_else(|_| ole.open_stream(&["PP97_DUALSTORAGE", "Current User"]))
            .ok();
        let encrypted = decrypt_powerpoint_document(
            &mut powerpoint_document,
            current_user_data.as_deref(),
            options.password,
        )?;

        // Parse document structure
        let mut parser = PptRecordParser::new();
        if let Some(encrypted) = &encrypted {
            parser.parse_document_at_offsets(&powerpoint_document, &encrypted.live_offsets)?;
        } else {
            parser.parse_document(&powerpoint_document)?;
        }

        // Build persist mapping for slide lookup (collect all records recursively)
        // Use zero-copy reference collection to avoid cloning all record data
        let all_records_ref = parser.find_records_ref();
        let mut persist_mapping = PersistMapping::build_from_records_ref(&all_records_ref);
        if let Some(encrypted) = &encrypted {
            persist_mapping = PersistMapping::new();
            for &(persist_id, offset) in &encrypted.mappings {
                persist_mapping.add_mapping(persist_id, offset);
            }
        }
        let current_user_data = current_user_data
            .as_deref()
            .ok_or_else(|| PptError::StreamNotFound("Current User".to_string()))?;
        let slide_directory =
            SlideDirectory::build(&powerpoint_document, current_user_data, &persist_mapping)?;

        // Try to read Pictures stream for image extraction
        #[cfg(feature = "imgconv")]
        let (pictures_data, blip_store) = if let Ok(mut pictures) = ole.open_stream(&["Pictures"]) {
            if let Some(encrypted) = &encrypted {
                decrypt_pictures(&mut pictures, &encrypted.crypto)?;
            }
            // Extract BLIP store from pictures data
            let store = ImageExtractor::extract_blip_store(&pictures)
                .ok()
                .map(|store| store.into_owned()); // Convert to 'static lifetime
            (Some(pictures), store)
        } else {
            (None, None)
        };

        Ok(Self {
            powerpoint_document,
            parser,
            persist_mapping,
            slide_directory,
            #[cfg(feature = "imgconv")]
            pictures_data,
            #[cfg(feature = "imgconv")]
            blip_store,
        })
    }

    /// Read the PowerPoint Document stream from OLE file.
    fn read_powerpoint_document<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Vec<u8>> {
        // Try primary location
        if let Ok(data) = ole.open_stream(&["PowerPoint Document"]) {
            return Ok(data);
        }

        // Try alternate location
        if let Ok(data) = ole.open_stream(&["PP97_DUALSTORAGE", "PowerPoint Document"]) {
            return Ok(data);
        }

        Err(PptError::InvalidFormat(
            "PowerPoint Document stream not found".to_string(),
        ))
    }

    /// Get iterator over all slides with zero-copy borrowing.
    ///
    /// # Performance
    ///
    /// - Returns lazy iterator (slides parsed on iteration)
    /// - Zero-copy: slides borrow from document data
    /// - Each slide lazily loads its shapes
    pub fn slides(&self) -> Result<Vec<Slide<'_>>> {
        let factory = SlideFactory::new(
            &self.powerpoint_document,
            &self.persist_mapping,
            &self.slide_directory,
        );

        factory
            .slides()
            .enumerate()
            .map(|(idx, slide_result)| {
                slide_result.map(|slide_data| Slide::from_slide_data(slide_data, idx + 1))
            })
            .collect()
    }

    /// Get the number of slides (actual Slide records only).
    #[inline]
    pub fn slide_count(&self) -> usize {
        self.slide_directory.len()
    }

    /// Return the validated logical slide directory.
    pub fn slide_directory(&self) -> &SlideDirectory {
        &self.slide_directory
    }

    /// Parse PowerPoint 12 round-trip metadata for every main master slide.
    ///
    /// The returned values follow the main-master order in the PowerPoint document stream.
    /// Embedded ECMA-376 packages are validated but remain inert; no external resources or
    /// executable content are activated.
    pub fn powerpoint12_main_master_metadata(&self) -> Result<Vec<PowerPoint12MainMasterMetadata>> {
        self.parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::MainMaster)
            .map(PowerPoint12MainMasterMetadata::parse)
            .collect()
    }

    /// Parse document-level PowerPoint 12 settings and round-trip metadata.
    pub fn powerpoint12_document_properties(&self) -> Result<PowerPoint12DocumentProperties> {
        let mut documents = self
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        PowerPoint12DocumentProperties::parse(document)
    }

    /// Return the document routing-slip metadata, when present.
    ///
    /// Routing metadata is inert. This method never contacts recipients, starts
    /// a mail client, or updates routing status.
    pub fn routing_slip(&self) -> Result<Option<PowerPointRoutingSlip>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let mut records = document
            .children
            .iter()
            .filter(|record| record.record_type == PptRecordType::DocRoutingSlipAtom);
        let Some(record) = records.next() else {
            return Ok(None);
        };
        if records.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple routing slips".to_string(),
            ));
        }
        PowerPointRoutingSlip::parse(record).map(Some)
    }

    /// Return the strictly validated document-level embedded sound collection.
    ///
    /// Sound bytes remain borrowed and inert. This method never decodes audio,
    /// invokes a codec, plays media, or resolves an external resource.
    pub fn embedded_sounds(&self) -> Result<Option<PowerPointSoundCollection<'_>>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let mut collections = document
            .children
            .iter()
            .filter(|record| record.record_type == PptRecordType::SoundCollection);
        let Some(collection) = collections.next() else {
            return Ok(None);
        };
        if collections.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple SoundCollection containers".to_string(),
            ));
        }
        PowerPointSoundCollection::parse(collection).map(Some)
    }

    /// Return strictly validated inert audio/video metadata from `ExObjListContainer`.
    ///
    /// Paths are never accessed and embedded sound bytes are never decoded or played.
    pub fn external_media(&self) -> Result<Option<PowerPointExternalMediaCollection>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let Some(media) = PowerPointExternalMediaCollection::parse(document)? else {
            return Ok(None);
        };
        let _ = PowerPointOleObjectCollection::parse(document)?;
        let sounds = self.embedded_sounds()?;
        media.validate_sound_collection(sounds.as_ref())?;
        Ok(Some(media))
    }

    /// Return inert embedded and linked OLE metadata without loading object storage.
    pub fn ole_objects(&self) -> Result<Option<PowerPointOleObjectCollection>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let Some(objects) = PowerPointOleObjectCollection::parse(document)? else {
            return Ok(None);
        };
        objects.validate_persist_mapping(&self.persist_mapping)?;
        Ok(Some(objects))
    }

    /// Resolve a persisted OLE, VBA, or ActiveX storage record as bounded opaque bytes.
    ///
    /// The returned payload is never decompressed, opened as an OLE filesystem, or executed.
    pub fn ole_storage(&self, persist_id: u32) -> Result<Option<PowerPointOleStorage>> {
        let Some(offset) = self.persist_mapping.get_offset(persist_id) else {
            return Ok(None);
        };
        let offset = usize::try_from(offset)
            .map_err(|_| PptError::Corrupted("OLE storage offset exceeds usize".to_string()))?;
        let (record, _) = PptRecord::parse(&self.powerpoint_document, offset)?;
        if record.record_type != PptRecordType::ExternalOleObjectStg {
            return Err(PptError::Corrupted(format!(
                "persist ID {persist_id} does not reference ExOleObjStg"
            )));
        }
        PowerPointOleStorage::parse(&record).map(Some)
    }

    /// Enumerate typed, inert native charts embedded as OLE objects.
    ///
    /// Embedded objects whose subtype or ProgID identifies `MSGraph.Chart` or
    /// `Excel.Chart` ([MS-PPT] 2.13.11) have their `ExOleObjStg` payload opened
    /// as a compound storage and their BIFF8 chart substreams parsed. Linked
    /// charts are never opened. Everything is inert: no formula evaluation, no
    /// rendering, and no OLE activation. A corrupt payload degrades to a
    /// per-object failure entry and never aborts the remaining charts.
    pub fn charts(&self) -> Result<crate::ppt::PowerPointChartInventory> {
        self.charts_with_limits(crate::xls::XlsChartLimits::default())
    }

    /// Enumerate native charts with caller-supplied BIFF8 chart resource limits.
    pub fn charts_with_limits(
        &self,
        limits: crate::xls::XlsChartLimits,
    ) -> Result<crate::ppt::PowerPointChartInventory> {
        crate::ppt::chart::enumerate(self, limits)
    }

    /// Resolve the live `DocumentContainer` record via the persist directory.
    ///
    /// Incrementally saved presentations can hold several `DocumentContainer`
    /// records; only the one referenced by the current `UserEditAtom` is live.
    pub(crate) fn live_document_record(&self) -> Result<PptRecord> {
        let persist_id = self.slide_directory.document_persist_id();
        let offset = self.persist_mapping.get_offset(persist_id).ok_or_else(|| {
            PptError::Corrupted(format!("document persist ID {persist_id} has no mapping"))
        })?;
        let offset = usize::try_from(offset)
            .map_err(|_| PptError::Corrupted("document offset exceeds usize".to_string()))?;
        let (record, _) = PptRecord::parse(&self.powerpoint_document, offset)?;
        if record.record_type != PptRecordType::Document {
            return Err(PptError::Corrupted(
                "document persist ID does not resolve to a DocumentContainer".to_string(),
            ));
        }
        Ok(record)
    }

    /// Return a payload-free descriptor for the document's VBA project storage.
    ///
    /// Macro bytes are not returned, decompressed, interpreted, or executed.
    /// A non-null persist reference must resolve to a `VbaProjectStg` record.
    pub fn vba_project_storage(&self) -> Result<Option<crate::ppt::PowerPointVbaProjectStorage>> {
        let records = self.parser.find_records_ref();
        let Some(info) = crate::ppt::PowerPointVbaInfo::parse_records(&records)? else {
            return Ok(None);
        };
        let storage = if info.persist_id_ref == 0 {
            None
        } else {
            let Some(storage) = self.ole_storage(info.persist_id_ref)? else {
                return Err(PptError::Corrupted(format!(
                    "VBAInfoAtom persist ID {} has no storage record",
                    info.persist_id_ref
                )));
            };
            Some(storage)
        };
        crate::ppt::PowerPointVbaProjectStorage::from_info_and_storage(info, storage.as_ref())
            .map(Some)
    }

    /// Return validated inert `VBAInfoAtom` metadata for the document.
    ///
    /// For richer outer-storage metadata without exposing VBA payload bytes,
    /// use [`Self::vba_project_storage`].
    pub fn vba_info(&self) -> Result<Option<crate::ppt::PowerPointVbaInfo>> {
        Ok(self.vba_project_storage()?.map(|storage| storage.info()))
    }

    /// Return the PPT10 modify-password metadata without verifying it.
    ///
    /// The returned value redacts its secret from `Debug`; callers must use an
    /// explicitly named accessor to inspect it. This method does not decrypt
    /// the presentation or grant modification access.
    pub fn modify_password(&self) -> Result<Option<crate::ppt::PowerPointModifyPassword>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointModifyPassword::parse_document(document)
    }

    /// Return the inert PowerPoint 10 privacy preference, if present.
    ///
    /// This accessor never removes document metadata or rewrites the file.
    pub fn privacy_settings(&self) -> Result<Option<crate::ppt::PowerPointPrivacySettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointPrivacySettings::parse_document(document)
    }

    /// Return typed PowerPoint 9 Presentation Advisor warning preferences.
    pub fn presentation_advisor_settings(
        &self,
    ) -> Result<Option<crate::ppt::PowerPointPresentationAdvisorSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointPresentationAdvisorSettings::parse_document(document)
    }

    /// Return typed PowerPoint 9 Web-document publishing preferences.
    ///
    /// This accessor never writes files, invokes a browser, or exports content.
    pub fn html_document_settings(
        &self,
    ) -> Result<Option<crate::ppt::PowerPointHtmlDocumentSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointHtmlDocumentSettings::parse_document(document)
    }

    /// Return inert PowerPoint 9 Web-publication metadata, if present.
    ///
    /// Named-show references are cross-validated against the presentation's
    /// named-show table. This accessor never writes files, resolves a URI,
    /// invokes a browser, or exports presentation content.
    pub fn html_publish_settings(
        &self,
    ) -> Result<Option<crate::ppt::PowerPointHtmlPublishSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointHtmlPublishSettings::parse_document(document)
    }

    /// Return all inert PowerPoint 9 presentation-broadcast descriptions.
    ///
    /// This accessor never contacts a server, opens a URL or ASD file, sends
    /// mail, records media, or starts a broadcast.
    pub fn broadcasts(&self) -> Result<crate::ppt::PowerPointBroadcasts> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointBroadcasts::parse_document(document)
    }

    /// Return the structurally decoded, inert PowerPoint 9 mail envelope.
    ///
    /// This accessor never sends mail, invokes a mail client, opens an
    /// attachment, or evaluates attachment bytes.
    pub fn envelope_data(&self) -> Result<Option<crate::ppt::PowerPointEnvelopeData>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointEnvelopeData::parse_document(document)
    }

    /// Validate and expose the specification-defined terminal document records.
    pub fn document_structure(&self) -> Result<crate::ppt::PowerPointDocumentStructure> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointDocumentStructure::parse(document)
    }

    /// Return inert PowerPoint 9 mail-envelope state, if present.
    ///
    /// This accessor never sends mail, invokes a mail client, or interprets the
    /// associated envelope payload.
    pub fn envelope_settings(&self) -> Result<Option<crate::ppt::PowerPointEnvelopeSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == PptRecordType::Document);
        let document = documents.next().ok_or_else(|| {
            PptError::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(PptError::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ppt::PowerPointEnvelopeSettings::parse_document(document)
    }

    /// Return strictly validated header/footer metadata from all specification-defined scopes.
    ///
    /// Date identifiers and text remain inert metadata. This method does not format dates,
    /// execute content, resolve resources, or modify the underlying OLE file.
    pub fn header_footers(&self) -> Result<PowerPointHeaderFooters> {
        let records = self.parser.find_records_ref();
        let mut values = PowerPointHeaderFooters::parse_record_tree(&records)?;

        let mut first_master_display = None;
        let mut master_ordinal = 0usize;
        for master in records
            .iter()
            .filter(|record| record.record_type == PptRecordType::MainMaster)
        {
            if let Some(display) = placeholder_display_from_record(master)? {
                let scope = PowerPointHeaderFooterScope::Local {
                    parent: PowerPointHeaderFooterParent::MainMaster,
                    parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(master_ordinal),
                };
                if values.has_scope(scope) {
                    values.attach_placeholder_display(scope, display.clone())?;
                }
                if first_master_display.is_none() {
                    first_master_display = Some(display);
                }
            }
            master_ordinal += 1;
        }
        let has_master_display = first_master_display.is_some();
        if let Some(display) = first_master_display {
            values.attach_placeholder_display(
                PowerPointHeaderFooterScope::PresentationSlides,
                display,
            )?;
        }

        let mut first_unoverridden_slide_display = None;
        let mut first_notes_display = None;
        for (ordinal, slide) in records
            .iter()
            .filter(|record| record.record_type == PptRecordType::Slide)
            .enumerate()
        {
            let scope = PowerPointHeaderFooterScope::Local {
                parent: PowerPointHeaderFooterParent::Slide,
                parent_ordinal: PowerPointHeaderFooterParentOrdinal::new(ordinal),
            };
            if let Some(display) = placeholder_display_from_record(slide)? {
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
                first_notes_display = placeholder_display_from_shapes(notes.shapes()?)?;
            }
        }
        if !has_master_display && let Some(display) = first_unoverridden_slide_display {
            values.attach_placeholder_display(
                PowerPointHeaderFooterScope::PresentationSlides,
                display,
            )?;
        }
        if let Some(display) = first_notes_display {
            values.attach_placeholder_display(
                PowerPointHeaderFooterScope::NotesAndHandouts,
                display,
            )?;
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
    pub fn text(&self) -> Result<String> {
        let slides = self.slides()?;
        let text_parts: Vec<String> = slides
            .iter()
            .filter_map(|slide| slide.text().ok().map(|s| s.to_string()))
            .filter(|text| !text.is_empty())
            .collect();

        Ok(if text_parts.is_empty() {
            String::from("No text content found in presentation")
        } else {
            text_parts.join("\n\n")
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
    /// Vector of (slide_number, text) tuples for each slide
    pub fn extract_text_fast(&self) -> Result<Vec<(usize, String)>> {
        let factory = SlideFactory::new(
            &self.powerpoint_document,
            &self.persist_mapping,
            &self.slide_directory,
        );

        let mut results = Vec::with_capacity(self.slide_directory.len());

        for (idx, slide_result) in factory.slides().enumerate() {
            let slide_data = slide_result?;

            // Pre-allocate string buffer
            let mut text = String::with_capacity(512);

            // Extract text from slide records without parsing shapes
            if let Ok(record_text) = slide_data.record.extract_text() {
                let trimmed = record_text.trim();
                if !trimmed.is_empty() {
                    text.push_str(trimmed);
                }
            }

            // Extract text from Escher/PPDrawing using the optimized path
            if let Some(ppdrawing) = slide_data
                .record
                .find_child(crate::consts::PptRecordType::PPDrawing)
                && let Ok(escher_text) = super::escher::extract_text_from_escher(&ppdrawing.data)
            {
                let trimmed = escher_text.trim();
                if !trimmed.is_empty() {
                    if !text.is_empty() {
                        text.push('\n');
                    }
                    text.push_str(trimmed);
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
    /// This extracts all embedded images from the Pictures stream.
    ///
    /// # Returns
    /// Vector of all extracted images with metadata
    ///
    /// # Example
    /// ```no_run
    /// use litchi_ole::ppt::Package;
    ///
    /// let mut pkg = Package::open("presentation.ppt")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for image in pres.extract_all_images()? {
    ///     let png_data = image.to_png(None, None)?;
    ///     std::fs::write(image.suggested_filename(), png_data)?;
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[cfg(feature = "imgconv")]
    pub fn extract_all_images(&self) -> Result<Vec<ExtractedImage<'static>>> {
        if let Some(ref pictures_data) = self.pictures_data {
            ImageExtractor::extract_from_pictures_stream(pictures_data)
                .map_err(|e| PptError::Corrupted(format!("Failed to extract images: {}", e)))
        } else {
            Ok(Vec::new())
        }
    }

    /// Extract an image by BLIP ID
    ///
    /// This method is used internally by PictureShape to resolve
    /// BLIP ID references to actual image data.
    ///
    /// # Arguments
    /// * `blip_id` - The BLIP ID from the shape's Escher properties
    ///
    /// # Returns
    /// The extracted image, or None if not found
    #[cfg(feature = "imgconv")]
    pub(crate) fn extract_image_by_blip_id(
        &self,
        blip_id: u32,
    ) -> Result<Option<ExtractedImage<'static>>> {
        // Extract all images and find the one matching the BLIP ID
        let images = self.extract_all_images()?;

        // BLIP ID is 1-based index
        let index = (blip_id.saturating_sub(1)) as usize;

        Ok(images.into_iter().nth(index))
    }

    /// Get the BLIP store (image metadata index)
    ///
    /// This provides access to image metadata without extracting the full image data.
    #[cfg(feature = "imgconv")]
    pub fn blip_store(&self) -> Option<&BlipStore<'static>> {
        self.blip_store.as_ref()
    }

    /// Check if the presentation has a Pictures stream
    #[cfg(feature = "imgconv")]
    pub fn has_pictures(&self) -> bool {
        self.pictures_data.is_some()
    }

    /// Strictly parse slide and notes editing-view information.
    pub fn slide_view_information(&self) -> Result<PowerPointSlideViewInformation> {
        let records = self.parser.find_records_ref();
        PowerPointSlideViewInformation::parse_records(&records)
    }

    /// Strictly parse outline and slide-sorter editing-view information.
    pub fn outline_sorter_view_information(
        &self,
    ) -> Result<PowerPointOutlineSorterViewInformation> {
        let records = self.parser.find_records_ref();
        PowerPointOutlineSorterViewInformation::parse_records(&records)
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

    /// Return all shape-scoped programmable tags in slide and shape order.
    pub fn shape_programmable_tags(
        &self,
    ) -> Result<Vec<crate::ppt::PowerPointPresentationShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(
            crate::ppt::PowerPointShapeProgrammableTagLimits::default(),
        )
    }

    /// Return the inert document-wide PowerPoint 11 smart-tag store.
    pub fn smart_tags(&self) -> Result<Option<crate::ppt::PowerPointSmartTagStore>> {
        let document = self.live_document_record()?;
        crate::ppt::PowerPointSmartTagStore::parse(&document)
    }

    /// Return all shape programmable tags with caller-supplied resource limits.
    pub fn shape_programmable_tags_with_limits(
        &self,
        limits: crate::ppt::PowerPointShapeProgrammableTagLimits,
    ) -> Result<Vec<crate::ppt::PowerPointPresentationShapeProgrammableTagsEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.shape_programmable_tags_with_limits(limits)? {
                result.push(
                    crate::ppt::PowerPointPresentationShapeProgrammableTagsEntry {
                        slide_number: slide.slide_number(),
                        shape_id: entry.shape_id,
                        programmable_tags: entry.programmable_tags,
                    },
                );
            }
        }
        Ok(result)
    }

    /// Return every typed shape-flag projection in slide and shape order.
    pub fn shape_flags(&self) -> Result<Vec<crate::ppt::PowerPointPresentationShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ppt::PowerPointShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    pub fn shape_flags_with_limits(
        &self,
        limits: crate::ppt::PowerPointShapeFlagLimits,
    ) -> Result<Vec<crate::ppt::PowerPointPresentationShapeFlagEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.shape_flags_with_limits(limits)? {
                result.push(crate::ppt::PowerPointPresentationShapeFlagEntry {
                    slide_number: slide.slide_number(),
                    shape_id: entry.shape_id,
                    projection: entry.projection,
                });
            }
        }
        Ok(result)
    }

    /// Return every context-validated presentation-slide placeholder.
    pub fn placeholder_atoms(
        &self,
    ) -> Result<Vec<crate::ppt::PowerPointPresentationPlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::ppt::PowerPointPlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    pub fn placeholder_atoms_with_limits(
        &self,
        limits: crate::ppt::PowerPointPlaceholderLimits,
    ) -> Result<Vec<crate::ppt::PowerPointPresentationPlaceholderEntry>> {
        let mut result = Vec::new();
        for slide in self.slides()? {
            for entry in slide.placeholder_atoms_with_limits(limits)? {
                result.push(crate::ppt::PowerPointPresentationPlaceholderEntry {
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
    /// CString names and `NamedShowSlides` (type=1042) slide ID arrays.
    ///
    /// # Returns
    ///
    /// A vector of `(name, slide_indices)` tuples for each custom show.
    /// Slide indices are 0-based.
    pub fn custom_shows(&self) -> Vec<ParsedCustomShow> {
        let mut shows = Vec::new();

        // Parse Document record from the stream
        let records = self.parser.find_records_ref();
        for record in &records {
            if record.record_type == PptRecordType::Document {
                // Find NamedShows container in Document
                for child in &record.children {
                    if child.record_type == PptRecordType::NamedShows {
                        Self::parse_named_shows(child, &mut shows);
                    }
                }
            }
        }

        shows
    }

    /// Parse NamedShow containers from a NamedShows container.
    fn parse_named_shows(
        named_shows: &super::records::PptRecord,
        shows: &mut Vec<ParsedCustomShow>,
    ) {
        for child in &named_shows.children {
            if child.record_type == PptRecordType::NamedShow {
                let mut name = String::new();
                let mut slide_indices = Vec::new();

                for sub in &child.children {
                    match sub.record_type {
                        PptRecordType::CString => {
                            // UTF-16LE name
                            let chars: Vec<u16> = sub
                                .data
                                .chunks_exact(2)
                                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                                .collect();
                            name = String::from_utf16_lossy(&chars);
                        },
                        PptRecordType::NamedShowSlides => {
                            // Array of u32 slide IDs (0x100 + slide_index)
                            for chunk in sub.data.chunks_exact(4) {
                                let slide_id =
                                    u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
                                // Convert slide ID (0x100+index) back to 0-based index
                                let index = slide_id.saturating_sub(0x100) as usize;
                                slide_indices.push(index);
                            }
                        },
                        _ => {},
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

fn placeholder_display_from_record(
    record: &crate::ppt::records::PptRecord,
) -> Result<Option<PowerPointHeaderFooterDisplayText>> {
    let Some(drawing) = record.find_child(PptRecordType::PPDrawing) else {
        return Ok(None);
    };
    let parsed =
        crate::ppt::escher::EscherShapeFactory::extract_shapes_from_drawing(&drawing.data)?;
    let mut shapes = Vec::with_capacity(parsed.len());
    for shape in &parsed {
        if let Some(shape) = Slide::<'static>::convert_escher_to_shape_enum(shape)? {
            shapes.push(shape);
        }
    }
    placeholder_display_from_shapes(&shapes)
}

fn placeholder_display_from_shapes(
    shapes: &[crate::ppt::shapes::ShapeEnum<'static>],
) -> Result<Option<PowerPointHeaderFooterDisplayText>> {
    use crate::ppt::shapes::PlaceholderType;

    let mut display = PowerPointHeaderFooterDisplayText::default();
    for shape in shapes {
        let Some(placeholder) = shape.as_placeholder() else {
            continue;
        };
        let target = match placeholder.placeholder_type() {
            PlaceholderType::DateAndTime => &mut display.user_date,
            PlaceholderType::Header => &mut display.header,
            PlaceholderType::Footer => &mut display.footer,
            _ => continue,
        };
        if target.is_some() {
            continue;
        }
        let text = shape.text()?;
        if !text.is_empty() && text != "*" {
            *target = Some(text);
        }
    }
    if display == PowerPointHeaderFooterDisplayText::default() {
        Ok(None)
    } else {
        Ok(Some(display))
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::records::PptRecord;

    fn record(record_type: PptRecordType, data: Vec<u8>, children: Vec<PptRecord>) -> PptRecord {
        PptRecord {
            record_type,
            record_type_raw: 0,
            version: 0,
            instance: 0,
            data_length: data.len() as u32,
            data,
            children,
        }
    }

    fn record_bytes(
        version: u16,
        instance: u16,
        record_type: PptRecordType,
        data: &[u8],
    ) -> Vec<u8> {
        let mut bytes = Vec::with_capacity(8 + data.len());
        bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
        bytes.extend_from_slice(&record_type.as_u16().to_le_bytes());
        bytes.extend_from_slice(&(data.len() as u32).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn presentation_with_vba_storage() -> Presentation {
        let mut atom_data = Vec::new();
        atom_data.extend_from_slice(&41u32.to_le_bytes());
        atom_data.extend_from_slice(&1u32.to_le_bytes());
        atom_data.extend_from_slice(&2u32.to_le_bytes());
        let atom = record_bytes(2, 0, PptRecordType::VBAInfoAtom, &atom_data);
        let vba_info = record_bytes(0x0f, 1, PptRecordType::VBAInfo, &atom);

        let mut storage_data = Vec::new();
        storage_data.extend_from_slice(&4096u32.to_le_bytes());
        storage_data.extend_from_slice(&[0x78, 0x9c, 1, 2, 3]);
        let storage = record_bytes(0, 3, PptRecordType::ExternalOleObjectStg, &storage_data);
        let storage_offset = vba_info.len() as u32;

        let mut powerpoint_document = vba_info;
        powerpoint_document.extend_from_slice(&storage);
        let mut parser = PptRecordParser::new();
        parser.parse_document(&powerpoint_document).unwrap();
        let mut persist_mapping = PersistMapping::new();
        persist_mapping.add_mapping(41, storage_offset);

        Presentation {
            powerpoint_document,
            parser,
            persist_mapping,
            slide_directory: SlideDirectory::new_for_test(0),
            #[cfg(feature = "imgconv")]
            pictures_data: None,
            #[cfg(feature = "imgconv")]
            blip_store: None,
        }
    }

    fn named_shows(children: Vec<PptRecord>) -> PptRecord {
        record(PptRecordType::NamedShows, Vec::new(), children)
    }

    fn named_show(name: &str, slide_ids: &[u32]) -> PptRecord {
        let name_bytes: Vec<u8> = name
            .encode_utf16()
            .flat_map(|unit| unit.to_le_bytes())
            .collect();
        let slide_bytes: Vec<u8> = slide_ids.iter().flat_map(|id| id.to_le_bytes()).collect();
        record(
            PptRecordType::NamedShow,
            Vec::new(),
            vec![
                record(PptRecordType::CString, name_bytes, Vec::new()),
                record(PptRecordType::NamedShowSlides, slide_bytes, Vec::new()),
            ],
        )
    }

    #[test]
    fn parses_named_shows_container() {
        let container = named_shows(vec![
            named_show("Demo Show", &[0x101, 0x103]),
            named_show("Short", &[0x100]),
        ]);

        let mut shows = Vec::new();
        Presentation::parse_named_shows(&container, &mut shows);

        assert_eq!(shows.len(), 2);
        assert_eq!(shows[0].name, "Demo Show");
        assert_eq!(shows[0].slide_indices, vec![1, 3]);
        assert_eq!(shows[1].name, "Short");
        assert_eq!(shows[1].slide_indices, vec![0]);
    }

    #[test]
    fn ignores_trailing_partial_slide_id_bytes() {
        let mut show = named_show("Odd", &[0x102]);
        // Append 3 stray bytes to the NamedShowSlides atom.
        show.children[1].data.extend_from_slice(&[0xAA, 0xBB, 0xCC]);
        let container = named_shows(vec![show]);

        let mut shows = Vec::new();
        Presentation::parse_named_shows(&container, &mut shows);

        assert_eq!(shows.len(), 1);
        assert_eq!(shows[0].slide_indices, vec![2]);
    }

    #[test]
    fn skips_named_show_without_name() {
        let show = record(
            PptRecordType::NamedShow,
            Vec::new(),
            vec![record(
                PptRecordType::NamedShowSlides,
                0x101u32.to_le_bytes().to_vec(),
                Vec::new(),
            )],
        );
        let container = named_shows(vec![show]);

        let mut shows = Vec::new();
        Presentation::parse_named_shows(&container, &mut shows);
        assert!(shows.is_empty());
    }

    #[test]
    fn vba_project_storage_returns_only_outer_metadata() {
        let presentation = presentation_with_vba_storage();

        let storage = presentation.vba_project_storage().unwrap().unwrap();
        assert_eq!(storage.persist_id_ref(), 41);
        assert!(storage.has_macros());
        assert!(storage.has_persisted_storage());
        assert_eq!(storage.stored_payload_len(), Some(5));
        assert_eq!(storage.declared_uncompressed_len(), Some(4096));
        assert_eq!(
            storage.compression(),
            Some(crate::ppt::PowerPointOleStorageCompression::Zlib {
                uncompressed_len: 4096,
            })
        );
        assert!(storage.may_contain_macro_code());
        assert_eq!(presentation.vba_info().unwrap(), Some(storage.info()));
    }
}
