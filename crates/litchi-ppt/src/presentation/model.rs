use crate::consts::RecordType;
use crate::document_properties::DocumentProperties12;
use crate::embedded::object::Collection as OleCollection;
#[cfg(feature = "vba-inspection")]
use crate::embedded::storage::Ref as StorageRef;
use crate::embedded::storage::{Kind as StorageKind, Storage};
use crate::external_media::Collection as MediaCollection;
use crate::header_footer::{
    HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope, HeaderFooters,
};
use crate::hyperlink::Hyperlinks;
use crate::main_master::MainMasterMetadata12;
use crate::non_zoom_view::OutlineSorterViewInformation;
/// High-performance Presentation API with zero-copy slide parsing.
use crate::package::{Error, RecordLimits, Result};
use crate::parsers::RecordParser;
use crate::persist::PersistMapping;
use crate::records::Record;
use crate::routing_slip::Slip;
use crate::slide::{ParsedComment, Slide, SlideDirectory, SlideFactory};
use crate::sound_collection::Collection;
use crate::view_info::SlideViewInformation;
#[cfg(feature = "vba-inspection")]
use litchi_cfb::OleFile;
use litchi_odraw::image::{File as ImageFile, Id as ImageId, Store as ImageStore};
#[cfg(feature = "vba-inspection")]
use std::io::Cursor;

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
pub struct Presentation {
    /// The main document stream data (owned for lifetime management)
    pub(super) powerpoint_document: Vec<u8>,
    /// Parsed record structure (reserved for future advanced parsing)
    #[allow(dead_code)]
    pub(crate) parser: RecordParser,
    /// Persist ID to offset mapping
    pub(crate) persist_mapping: PersistMapping,
    pub(super) slide_directory: SlideDirectory,
    /// Pictures stream data (for image extraction)
    pub(super) pictures_data: Option<Vec<u8>>,
    /// Record limits retained for lazy live-persist parsing.
    pub(crate) record_limits: RecordLimits,
}

impl Presentation {
    /// Get iterator over all slides with zero-copy borrowing.
    ///
    /// # Performance
    ///
    /// - Returns lazy iterator (slides parsed on iteration)
    /// - Zero-copy: slides borrow from document data
    /// - Each slide lazily loads its shapes
    pub fn slides(&self) -> Result<Vec<Slide<'_>>> {
        let factory = SlideFactory::new_with_limits(
            &self.powerpoint_document,
            &self.persist_mapping,
            &self.slide_directory,
            self.record_limits,
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
    pub fn powerpoint12_main_master_metadata(&self) -> Result<Vec<MainMasterMetadata12>> {
        self.parser
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
    pub fn main_master_programmable_tags(&self) -> Result<Vec<crate::ProgTags>> {
        self.main_master_programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return main-master programmable tags with caller-supplied resource limits.
    pub fn main_master_programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Vec<crate::ProgTags>> {
        self.parser
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
    pub fn text_metachars(&self) -> Result<Vec<crate::text_metachar::TextMetachar>> {
        use crate::text_metachar::TextMetachar;

        fn collect(record: &Record, out: &mut Vec<TextMetachar>) -> Result<()> {
            if record.record_type == RecordType::PPDrawing {
                for textbox in super::codec::drawing_textboxes(&record.data)? {
                    let wrapper = crate::EscherTextboxWrapper::new(textbox.data().to_vec())?;
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
        for record in self.parser.find_records_ref() {
            collect(record, &mut result)?;
        }
        Ok(result)
    }

    /// The normal three-pane view's splitter state (`NormalViewSetInfo9`,
    /// MS-PPT 2.4.21.2), when the document declares one. Files with multiple
    /// top-level Document containers yield the first occurrence.
    pub fn normal_view_set_info(&self) -> Result<Option<crate::view_set_info::NormalViewSet>> {
        for record in self.parser.find_records_ref() {
            if record.record_type == RecordType::NormalViewSetInfo9 {
                return crate::view_set_info::NormalViewSet::parse_record(record).map(Some);
            }
        }
        Ok(None)
    }

    /// The notes-text view's scaling state (`NotesTextViewInfo9`, MS-PPT
    /// 2.4.21.4), when the document declares one. Files with multiple
    /// top-level Document containers yield the first occurrence.
    pub fn notes_text_view_info(&self) -> Result<Option<crate::view_set_info::NotesTextViewInfo>> {
        for record in self.parser.find_records_ref() {
            if record.record_type == RecordType::NotesTextViewInfo9 {
                return crate::view_set_info::NotesTextViewInfo::parse_record(record).map(Some);
            }
        }
        Ok(None)
    }

    /// The default language and spelling settings for document text
    /// (`TextSIExceptionAtom`, MS-PPT 2.9.31), when declared.
    pub fn text_special_info_defaults(
        &self,
    ) -> Result<Option<crate::text_si_exception::TextSpecialInfoDefaults>> {
        for record in self.parser.find_records_ref() {
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
    pub fn outline_text_refs(&self) -> Result<Vec<crate::OutlineTextRef>> {
        use crate::OutlineTextRef;

        fn collect(record: &Record, out: &mut Vec<OutlineTextRef>) -> Result<()> {
            if record.record_type == RecordType::PPDrawing {
                for textbox in super::codec::drawing_textboxes(&record.data)? {
                    let wrapper = crate::EscherTextboxWrapper::new(textbox.data().to_vec())?;
                    out.extend_from_slice(wrapper.outline_text_refs());
                }
            } else if record.record_type == RecordType::OutlineTextRefAtom {
                out.push(OutlineTextRef::parse_record(record)?);
            }
            Ok(())
        }

        let mut result = Vec::new();
        for record in self.parser.find_records_ref() {
            collect(record, &mut result)?;
        }
        Ok(result)
    }

    /// Parse document-level PowerPoint 12 settings and round-trip metadata.
    pub fn powerpoint12_document_properties(&self) -> Result<DocumentProperties12> {
        let mut documents = self
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        DocumentProperties12::parse(document)
    }

    /// Return the document routing-slip metadata, when present.
    ///
    /// Routing metadata is inert. This method never contacts recipients, starts
    /// a mail client, or updates routing status.
    pub fn routing_slip(&self) -> Result<Option<Slip>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let mut records = document
            .children
            .iter()
            .filter(|record| record.record_type == RecordType::DocRoutingSlipAtom);
        let Some(record) = records.next() else {
            return Ok(None);
        };
        if records.next().is_some() {
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
    pub fn embedded_sounds(&self) -> Result<Option<Collection<'_>>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
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
    /// every non-null reference must resolve in the document SoundCollection.
    pub fn validate_interaction_sound_references(&self) -> Result<()> {
        let sounds = self.embedded_sounds()?;
        let validate = |interaction: &crate::Interaction| -> Result<()> {
            if interaction.sound_id == 0 {
                return Ok(());
            }
            let sounds = sounds.as_ref().ok_or_else(|| {
                Error::Corrupted(
                    "interaction references a sound but the document has no SoundCollection"
                        .to_string(),
                )
            })?;
            interaction.validate_sound_collection(sounds)
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
    pub fn external_media(&self) -> Result<Option<MediaCollection>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let Some(media) = MediaCollection::parse(document)? else {
            return Ok(None);
        };
        let _ = OleCollection::parse(document)?;
        let sounds = self.embedded_sounds()?;
        media.validate_sound_collection(sounds.as_ref())?;
        Ok(Some(media))
    }

    /// Return the document's strictly validated hyperlink definitions.
    ///
    /// Targets remain inert: this method never opens a URL, path, presentation,
    /// or named show.
    pub fn hyperlinks(&self) -> Result<Hyperlinks> {
        Hyperlinks::parse(&self.live_document_record()?)
    }

    /// Return inert embedded and linked OLE metadata without loading object storage.
    pub fn ole_objects(&self) -> Result<Option<OleCollection>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        let Some(objects) = OleCollection::parse(document)? else {
            return Ok(None);
        };
        objects.validate_persist_mapping(&self.persist_mapping)?;
        Ok(Some(objects))
    }

    /// Resolve a persisted embedded OLE-object storage as bounded inert bytes.
    pub fn ole_storage(&self, persist_id: u32) -> Result<Option<Storage>> {
        self.ole_storage_as(persist_id, StorageKind::OleObject)
    }

    /// Resolve a persisted storage with the kind supplied by its referencing record.
    pub fn ole_storage_as(&self, persist_id: u32, kind: StorageKind) -> Result<Option<Storage>> {
        let Some(offset) = self.persist_mapping.get_offset(persist_id) else {
            return Ok(None);
        };
        let offset = usize::try_from(offset)
            .map_err(|_| Error::Corrupted("OLE storage offset exceeds usize".to_string()))?;
        let (record, _) =
            Record::parse_with_limits(&self.powerpoint_document, offset, self.record_limits)?;
        if record.record_type != RecordType::ExternalOleObjectStg {
            return Err(Error::Corrupted(format!(
                "persist ID {persist_id} does not reference ExOleObjStg"
            )));
        }
        Storage::parse_as(&record, kind).map(Some)
    }

    /// Enumerate inert native charts as neutral Office Graph views.
    ///
    /// Embedded objects whose subtype or ProgID identifies `MSGraph.Chart` or
    /// `Excel.Chart` ([MS-PPT] 2.13.11) have their `ExOleObjStg` payload opened
    /// as a bounded compound storage and their chart substreams validated by
    /// `litchi-ograph`. Linked charts are never opened. A corrupt payload
    /// degrades to a per-object failure without aborting the remaining charts.
    pub fn charts(&self) -> Result<crate::chart::Inventory> {
        self.charts_with(litchi_ograph::Limits::default())
    }

    /// Enumerate native charts with explicit neutral resource limits.
    pub fn charts_with(&self, limits: litchi_ograph::Limits) -> Result<crate::chart::Inventory> {
        crate::chart::enumerate(self, limits)
    }

    /// Return the live document-comparison snapshot.
    ///
    /// Review metadata is inert: this accessor never compares presentations,
    /// opens the embedded reviewer document, or applies/rejects a change.
    pub fn document_comparison(&self) -> Result<crate::document_comparison::Snapshot> {
        crate::document_comparison::package::from_presentation(self)
    }

    /// Parse the base and PowerPoint 10 font owners from the exact live
    /// `DocumentContainer`. Embedded EOT payloads remain inert.
    pub fn fonts(&self) -> Result<crate::font::FontCollections> {
        self.fonts_with_limits(crate::font::Limits::default())
    }

    /// Parse live font owners with explicit composed record/font limits.
    pub fn fonts_with_limits(
        &self,
        mut limits: crate::font::Limits,
    ) -> Result<crate::font::FontCollections> {
        limits.records = limits.records.constrained_by(self.record_limits);
        crate::font::FontCollections::parse_with_limits(&self.live_document_record()?, limits)
    }

    /// Resolve the live `DocumentContainer` record via the persist directory.
    ///
    /// Incrementally saved presentations can hold several `DocumentContainer`
    /// records; only the one referenced by the current `UserEditAtom` is live.
    pub(crate) fn live_document_record(&self) -> Result<Record> {
        let persist_id = self.slide_directory.document_persist_id();
        let offset = self.persist_mapping.get_offset(persist_id).ok_or_else(|| {
            Error::Corrupted(format!("document persist ID {persist_id} has no mapping"))
        })?;
        let offset = usize::try_from(offset)
            .map_err(|_| Error::Corrupted("document offset exceeds usize".to_string()))?;
        let (record, _) =
            Record::parse_with_limits(&self.powerpoint_document, offset, self.record_limits)?;
        if record.record_type != RecordType::Document {
            return Err(Error::Corrupted(
                "document persist ID does not resolve to a DocumentContainer".to_string(),
            ));
        }
        Ok(record)
    }

    pub(crate) fn document_stream(&self) -> &[u8] {
        &self.powerpoint_document
    }

    /// Return a payload-free descriptor for the document's VBA project storage.
    ///
    /// Macro bytes are not returned, decompressed, interpreted, or executed.
    /// A non-null persist reference must resolve to a `VbaProjectStg` record.
    #[cfg(feature = "vba-inspection")]
    pub fn vba_project_storage(&self) -> Result<Option<crate::VbaProjectStorage>> {
        let records = self.parser.find_records_ref();
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
    #[cfg(feature = "vba-inspection")]
    pub fn vba_info(&self) -> Result<Option<crate::VbaInfo>> {
        Ok(self.vba_project_storage()?.map(|storage| storage.info()))
    }

    /// Parse the embedded MS-OVBA project with safe default limits.
    ///
    /// The outer zlib stream, inner CFB, and MS-OVBA compressed containers are
    /// bounded, and source is never compiled, interpreted, or executed. Use
    /// [`Self::vba_with`] to supply custom ceilings.
    #[cfg(feature = "vba-inspection")]
    pub fn vba(
        &self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, crate::VbaProjectError> {
        self.vba_with(&crate::VbaProjectLimits::default())
    }

    /// Parse the embedded MS-OVBA project with explicit resource limits.
    ///
    /// Stored and declared outer sizes are checked on a borrowed record view
    /// before any VBA payload is copied or decompressed.
    #[cfg(feature = "vba-inspection")]
    pub fn vba_with(
        &self,
        limits: &crate::VbaProjectLimits,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, crate::VbaProjectError> {
        let records = self.parser.find_records_ref();
        let Some(info) = crate::VbaInfo::parse_records(&records)? else {
            return Ok(None);
        };
        let storage = self.resolve_vba_storage(info)?;
        let summary = crate::VbaProjectStorage::from_info_and_metadata(
            info,
            storage.map(StorageRef::metadata),
        )?;
        if !summary.has_persisted_storage() {
            return Ok(None);
        }
        let Some(storage) = storage else {
            return Err(Error::Corrupted(format!(
                "VBAInfoAtom persist ID {} has no storage record",
                summary.persist_id_ref()
            ))
            .into());
        };
        let storage = storage.check_stored_limit(limits.max_stored_bytes)?;
        let cfb = storage.decompressed_bytes(limits.max_cfb_bytes)?;
        let mut ole = OleFile::open(Cursor::new(cfb.as_ref()))
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
            .persist_mapping
            .get_offset(info.persist_id_ref)
            .ok_or_else(|| {
                Error::Corrupted(format!(
                    "VBAInfoAtom persist ID {} has no storage record",
                    info.persist_id_ref
                ))
            })?;
        let offset = usize::try_from(offset)
            .map_err(|_| Error::Corrupted("VBA storage offset exceeds usize".to_string()))?;
        StorageRef::parse_at(&self.powerpoint_document, offset, StorageKind::VbaProject).map(Some)
    }

    /// Return the PPT10 modify-password metadata without verifying it.
    ///
    /// The returned value redacts its secret from `Debug`; callers must use an
    /// explicitly named accessor to inspect it. This method does not decrypt
    /// the presentation or grant modification access.
    pub fn modify_password(&self) -> Result<Option<crate::ModifyPassword>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::ModifyPassword::parse_document(document)
    }

    /// Return the inert PowerPoint 10 privacy preference, if present.
    ///
    /// This accessor never removes document metadata or rewrites the file.
    pub fn privacy_settings(&self) -> Result<Option<crate::PrivacySettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::PrivacySettings::parse_document(document)
    }

    /// Return typed PowerPoint 9 Presentation Advisor warning preferences.
    pub fn presentation_advisor_settings(
        &self,
    ) -> Result<Option<crate::PresentationAdvisorSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::PresentationAdvisorSettings::parse_document(document)
    }

    /// Return typed PowerPoint 9 Web-document publishing preferences.
    ///
    /// This accessor never writes files, invokes a browser, or exports content.
    pub fn html_document_settings(&self) -> Result<Option<crate::HtmlDocumentSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::HtmlDocumentSettings::parse_document(document)
    }

    /// Return inert PowerPoint 9 Web-publication metadata, if present.
    ///
    /// Named-show references are cross-validated against the presentation's
    /// named-show table. This accessor never writes files, resolves a URI,
    /// invokes a browser, or exports presentation content.
    pub fn html_publish_settings(&self) -> Result<Option<crate::HtmlPublishSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::HtmlPublishSettings::parse_document(document)
    }

    /// Return all inert PowerPoint 9 presentation-broadcast descriptions.
    ///
    /// This accessor never contacts a server, opens a URL or ASD file, sends
    /// mail, records media, or starts a broadcast.
    pub fn broadcasts(&self) -> Result<crate::Broadcasts> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::Broadcasts::parse_document(document)
    }

    /// Return the structurally decoded, inert PowerPoint 9 mail envelope.
    ///
    /// This accessor never sends mail, invokes a mail client, opens an
    /// attachment, or evaluates attachment bytes.
    pub fn envelope_data(&self) -> Result<Option<crate::EnvelopeData>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::EnvelopeData::parse_document(document)
    }

    /// Validate and expose the specification-defined terminal document records.
    pub fn document_structure(&self) -> Result<crate::DocumentStructure> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::DocumentStructure::parse(document)
    }

    /// Return the document-wide `DocumentAtom` (MS-PPT 2.4.2), if present.
    ///
    /// Slide geometry, OLE server zoom, master persist references, and display
    /// flags are inert metadata and nothing is rendered. Incremental histories
    /// are resolved through the live document persist mapping.
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
    pub fn color_schemes(&self) -> Result<Vec<crate::ColorSchemeAtom>> {
        let mut schemes = Vec::new();
        for record in self.parser.find_records_ref() {
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

    /// Return inert PowerPoint 9 mail-envelope state, if present.
    ///
    /// This accessor never sends mail, invokes a mail client, or interprets the
    /// associated envelope payload.
    pub fn envelope_settings(&self) -> Result<Option<crate::EnvelopeSettings>> {
        let records = self.parser.find_records_ref();
        let mut documents = records
            .into_iter()
            .filter(|record| record.record_type == RecordType::Document);
        let document = documents.next().ok_or_else(|| {
            Error::Corrupted("PowerPoint document has no Document container".to_string())
        })?;
        if documents.next().is_some() {
            return Err(Error::Corrupted(
                "PowerPoint document has multiple Document containers".to_string(),
            ));
        }
        crate::EnvelopeSettings::parse_document(document)
    }

    /// Return strictly validated header/footer metadata from all specification-defined scopes.
    ///
    /// Date identifiers and text remain inert metadata. This method does not format dates,
    /// execute content, resolve resources, or modify the underlying OLE file.
    pub fn header_footers(&self) -> Result<HeaderFooters> {
        let records = self.parser.find_records_ref();
        let mut values = HeaderFooters::parse_record_tree(&records)?;

        let mut first_master_display = None;
        for (master_ordinal, master) in records
            .iter()
            .filter(|record| record.record_type == RecordType::MainMaster)
            .enumerate()
        {
            if let Some(display) = super::codec::placeholder_display_from_record(master)? {
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
            if let Some(display) = super::codec::placeholder_display_from_record(slide)? {
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
        let factory = SlideFactory::new_with_limits(
            &self.powerpoint_document,
            &self.persist_mapping,
            &self.slide_directory,
            self.record_limits,
        );

        let mut results = Vec::with_capacity(self.slide_directory.len());

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

            // Extract text from Escher/PPDrawing using the optimized path
            if let Some(ppdrawing) = slide_data.record.find_child(RecordType::PPDrawing) {
                let escher_text = crate::odraw::text_from_drawing(&ppdrawing.data)?;
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
    /// Images are returned in semantic BStore order. FBSE metadata comes from
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
    pub fn images(&self) -> Result<Vec<ImageFile<'_>>> {
        if let Some(store) = self.image_store()? {
            return litchi_odraw::image::all(&store, self.pictures_data.as_deref())
                .map_err(Error::from);
        }
        self.pictures_data.as_deref().map_or_else(
            || Ok(Vec::new()),
            |pictures| litchi_odraw::image::delay(pictures).map_err(Error::from),
        )
    }

    /// Extract an image by BLIP ID
    ///
    /// This method is used internally by PictureShape to resolve
    /// BLIP ID references to actual image data.
    ///
    /// # Arguments
    /// * `id` - Checked BLIP ID from the shape's Escher properties
    ///
    /// # Returns
    /// The extracted image, or None if not found
    pub fn image(&self, id: ImageId) -> Result<Option<ImageFile<'_>>> {
        let Some(store) = self.image_store()? else {
            return Ok(None);
        };
        litchi_odraw::image::get(&store, id, self.pictures_data.as_deref()).map_err(Error::from)
    }

    /// Resolves a raw one-based host index after checking its OfficeArt range.
    pub fn image_at(&self, index: u32) -> Result<Option<ImageFile<'_>>> {
        self.image(ImageId::new(index)?)
    }

    /// Parses the borrowed BLIP store from the document's drawing group.
    ///
    /// No parsed store is retained inside `Presentation`; this avoids
    /// self-referential storage while keeping the view zero-copy.
    pub fn image_store(&self) -> Result<Option<ImageStore<'_>>> {
        let mut store = None;
        for record in self
            .parser
            .find_records_ref()
            .into_iter()
            .filter(|record| record.record_type == RecordType::PPDrawingGroup)
        {
            let candidate = litchi_odraw::image::store(&record.data).map_err(Error::from)?;
            if let Some(candidate) = candidate
                && store.replace(candidate).is_some()
            {
                return Err(Error::Corrupted(
                    "Presentation contains multiple OfficeArt BStore containers".to_string(),
                ));
            }
        }
        Ok(store)
    }

    /// Check if the presentation has a Pictures stream
    pub fn has_pictures(&self) -> bool {
        self.pictures_data.is_some()
    }

    /// Strictly parse slide and notes editing-view information.
    pub fn slide_view_information(&self) -> Result<SlideViewInformation> {
        let records = self.parser.find_records_ref();
        SlideViewInformation::parse_records(&records)
    }

    /// Strictly parse outline and slide-sorter editing-view information.
    pub fn outline_sorter_view_information(&self) -> Result<OutlineSorterViewInformation> {
        let records = self.parser.find_records_ref();
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
    pub fn comment_catalog(&self) -> Result<crate::comments::Catalog> {
        let authors = crate::comments::Authors::parse(&self.live_document_record()?)?;
        crate::comments::Catalog::from_parts(authors, self.comments()?)
    }

    /// Return all shape-scoped programmable tags in slide and shape order.
    pub fn shape_programmable_tags(
        &self,
    ) -> Result<Vec<crate::PresentationShapeProgrammableTagsEntry>> {
        self.shape_programmable_tags_with_limits(crate::ShapeProgrammableTagLimits::default())
    }

    /// Return the inert document-wide PowerPoint 11 smart-tag store.
    pub fn smart_tags(&self) -> Result<Option<crate::SmartTagStore>> {
        let document = self.live_document_record()?;
        crate::SmartTagStore::parse(&document)
    }

    /// Return the typed document-level programmable tags (MS-PPT 2.4.23), when
    /// the document carries a `DocProgTagsContainer`.
    ///
    /// Tag payloads are inert: they are parsed and preserved, never executed,
    /// loaded, or resolved. Use
    /// [`crate::ProgTags::document_extensions`] to decode the
    /// versioned binary-tag payloads into typed extension structs.
    pub fn programmable_tags(&self) -> Result<Option<crate::ProgTags>> {
        self.programmable_tags_with_limits(crate::ProgTagLimits::default())
    }

    /// Return document-level programmable tags with caller-supplied resource limits.
    pub fn programmable_tags_with_limits(
        &self,
        limits: crate::ProgTagLimits,
    ) -> Result<Option<crate::ProgTags>> {
        crate::ProgTags::parse_document(&self.live_document_record()?, limits)
    }

    /// Return all shape programmable tags with caller-supplied resource limits.
    pub fn shape_programmable_tags_with_limits(
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
    pub fn shape_flags(&self) -> Result<Vec<crate::PresentationShapeFlagEntry>> {
        self.shape_flags_with_limits(crate::ShapeFlagLimits::default())
    }

    /// Return shape flags with caller-supplied client-data resource limits.
    pub fn shape_flags_with_limits(
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
    pub fn placeholder_atoms(&self) -> Result<Vec<crate::PresentationPlaceholderEntry>> {
        self.placeholder_atoms_with_limits(crate::PlaceholderLimits::default())
    }

    /// Return placeholders with caller-supplied client-data limits.
    pub fn placeholder_atoms_with_limits(
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

    /// Parse NamedShow containers from a NamedShows container.
    pub(super) fn parse_named_shows(named_shows: &Record, shows: &mut Vec<ParsedCustomShow>) {
        for child in &named_shows.children {
            if child.record_type == RecordType::NamedShow {
                let mut name = String::new();
                let mut slide_indices = Vec::new();

                for sub in &child.children {
                    match sub.record_type {
                        RecordType::CString => {
                            // UTF-16LE name
                            let chars: Vec<u16> = sub
                                .data
                                .chunks_exact(2)
                                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                                .collect();
                            name = String::from_utf16_lossy(&chars);
                        },
                        RecordType::NamedShowSlides => {
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
