//! PPT OLE2 package and stream assembly.

use super::super::escher::{UserShapeData, create_dg_container_with_charts, create_dgg_container};
use super::super::master_drawing::build_master_ppdrawing;
use super::super::notes::{NotesContainerBuilder, NotesPage};
use super::super::persist::{PersistPtrBuilder, UserEditAtom};
use super::super::records::{
    RecordBuilder, create_document_atom_with_font_embedding, create_end_document,
    create_environment_minimal, create_environment_with_font_collection,
    create_main_master_container, create_slide_list_with_text_master, record_type,
    wrap_dg_into_ppdrawing, wrap_dgg_into_ppdrawing_group,
};
use super::super::spec::{BinaryTagData, ColorScheme, SlideLayoutType, Tag10, slide_flags};
use super::codec::{
    append_child_to_built_container, build_writer_sound_collection,
    convert_shape_to_escher_with_sound_mapping,
};
use super::model::{WriteError, Writer};
#[cfg(feature = "encryption")]
use crate::encryption::{encrypt_pictures_for_write, encrypt_powerpoint_document_for_write};
use litchi_cfb::writer::OleWriter;

pub(in crate::writer::core) struct FontPublicationPlan {
    pub(in crate::writer::core) environment: Vec<u8>,
    pub(in crate::writer::core) powerpoint10_records: Vec<Vec<u8>>,
    pub(in crate::writer::core) save_with_fonts: bool,
}

impl Writer {
    pub(in crate::writer::core) fn font_publication_plan(
        &self,
    ) -> Result<FontPublicationPlan, WriteError> {
        let base_record = self
            .fonts
            .base_record_bytes()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?
            .ok_or_else(|| {
                WriteError::InvalidData("PowerPoint base font collection is absent".to_string())
            })?;
        let environment = if self.has_historical_default_font_catalog() {
            create_environment_minimal()?
        } else {
            create_environment_with_font_collection(&base_record)?
        };
        let powerpoint10_records = self
            .fonts
            .powerpoint10_records()
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;
        Ok(FontPublicationPlan {
            environment,
            powerpoint10_records,
            save_with_fonts: self.fonts.has_embedded_fonts(),
        })
    }

    fn has_historical_default_font_catalog(&self) -> bool {
        if self.fonts.international.is_some() || self.fonts.embedding_flags.is_some() {
            return false;
        }
        let Some(base) = self.fonts.base.as_ref() else {
            return false;
        };
        let [font] = base.fonts.as_slice() else {
            return false;
        };
        font.index == 0
            && font.raw_instance == 0
            && font.name == "Arial"
            && font.charset == 0
            && font.font_flags == 0
            && !font.embedded_subset
            && font.font_type_flags == 0x04
            && !font.raster
            && !font.device
            && font.truetype
            && !font.no_substitution
            && font.pitch_and_family == 0x22
            && font.embedded_fonts.is_empty()
    }
}

#[cfg(test)]
mod font_plan_tests {
    use super::*;

    #[test]
    fn default_catalog_preserves_the_historical_environment_bytes() {
        let writer = Writer::new();
        let plan = writer.font_publication_plan().unwrap();
        assert_eq!(plan.environment, create_environment_minimal().unwrap());
        assert!(plan.powerpoint10_records.is_empty());
        assert!(!plan.save_with_fonts);
    }

    #[test]
    fn built_in_master_style_font_zero_is_validated_before_destinations_change() {
        let mut writer = Writer::new();
        writer.fonts.base = Some(crate::FontCollection::new(crate::FontScope::Base));

        let mut output = std::io::Cursor::new(Vec::new());
        assert!(writer.write_to(&mut output).is_err());
        assert!(output.get_ref().is_empty());

        let path = std::env::temp_dir().join(format!(
            "litchi-ppt-invalid-master-font-{}-{}.ppt",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        assert!(writer.save(&path).is_err());
        assert!(!path.exists());
    }
}

/// Build a minimal, valid Current User stream referencing the given UserEditAtom offset.
fn build_current_user_stream(offset_to_current_edit: u32, encrypted: bool) -> Vec<u8> {
    // Build per Apache POI CurrentUserAtom:
    // [0..3]   atomHeader = {0x00,0x00,0xF6,0x0F}
    // [4..7]   atomSize = 20 + 4 + lenAsciiUser (we use 0) => 24
    // [8..11]  details size = 20
    // [12..15] headerToken (unencrypted) = 0xE391C05F (bytes {95,-64,-111,-29})
    // [16..19] offsetToCurrentEdit
    // [20..21] lenUserName (ANSI) = 0
    // [22..23] docFinalVersion = 0x03F4
    // [24]     docMajorNo = 3
    // [25]     docMinorNo = 0
    // [26..27] reserved = 0
    // [28..31] releaseVersion = 8
    // [32..]   unicode username (2*len) (none)
    let mut s = Vec::with_capacity(32);
    // atomHeader
    s.extend_from_slice(&[0x00, 0x00, 0xF6, 0x0F]);
    // atomSize (20 + 4 + lenAsciiUsername)
    s.extend_from_slice(&24u32.to_le_bytes());
    // details size (20)
    s.extend_from_slice(&20u32.to_le_bytes());
    // headerToken
    let token: u32 = if encrypted { 0xF3D1_C4DF } else { 0xE391_C05F };
    s.extend_from_slice(&token.to_le_bytes());
    // current edit offset
    s.extend_from_slice(&offset_to_current_edit.to_le_bytes());
    // username length (ANSI)
    s.extend_from_slice(&0u16.to_le_bytes());
    // doc final version
    s.extend_from_slice(&0x03F4u16.to_le_bytes());
    // major/minor
    s.push(3u8);
    s.push(0u8);
    // reserved
    s.extend_from_slice(&[0u8; 2]);
    // release version
    s.extend_from_slice(&8u32.to_le_bytes());
    // no username
    s
}

fn build_summary_information_stream() -> Vec<u8> {
    let mut s = Vec::new();
    s.extend_from_slice(&0xFFFEu16.to_le_bytes());
    s.extend_from_slice(&0u16.to_le_bytes());
    s.extend_from_slice(&0u32.to_le_bytes());
    s.extend_from_slice(&[0u8; 16]);
    s.extend_from_slice(&1u32.to_le_bytes());
    let fmtid: [u8; 16] = [
        0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3,
        0xD9,
    ];
    s.extend_from_slice(&fmtid);
    let section_offset = 48u32;
    s.extend_from_slice(&section_offset.to_le_bytes());
    let mut section = Vec::new();
    section.extend_from_slice(&0u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&16u32.to_le_bytes());
    section.extend_from_slice(&2u16.to_le_bytes());
    section.extend_from_slice(&0u16.to_le_bytes());
    section.extend_from_slice(&(1252i16).to_le_bytes());
    section.extend_from_slice(&0i16.to_le_bytes());
    let size = section.len() as u32;
    section[0..4].copy_from_slice(&size.to_le_bytes());
    s.extend_from_slice(&section);
    s
}

fn build_document_summary_information_stream() -> Vec<u8> {
    let mut s = Vec::new();
    s.extend_from_slice(&0xFFFEu16.to_le_bytes());
    s.extend_from_slice(&0u16.to_le_bytes());
    s.extend_from_slice(&0u32.to_le_bytes());
    s.extend_from_slice(&[0u8; 16]);
    s.extend_from_slice(&1u32.to_le_bytes());
    let fmtid: [u8; 16] = [
        0x02, 0xD5, 0xCD, 0xD5, 0x9C, 0x2E, 0x1B, 0x10, 0x93, 0x97, 0x08, 0x00, 0x2B, 0x2C, 0xF9,
        0xAE,
    ];
    s.extend_from_slice(&fmtid);
    let section_offset = 48u32;
    s.extend_from_slice(&section_offset.to_le_bytes());
    let mut section = Vec::new();
    section.extend_from_slice(&0u32.to_le_bytes());
    section.extend_from_slice(&0u32.to_le_bytes());
    let size = section.len() as u32;
    section[0..4].copy_from_slice(&size.to_le_bytes());
    s.extend_from_slice(&section);
    s
}

impl Writer {
    /// Save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), WriteError>` - Success or error
    ///
    /// # Implementation
    ///
    /// This generates a complete PowerPoint 97-2003 binary file conforming to MS-PPT specification:
    /// - PPT record structures - [MS-PPT] Section 2.3
    /// - Escher drawing containers - [MS-ODRAW] Section 2.2
    /// - PersistPtr directory - [MS-PPT] Section 2.4.16
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), WriteError> {
        self.validate_encryption()?;
        self.validate_references()?;
        let font_plan = self.font_publication_plan()?;
        let modify_password_atom = self.build_modify_password_atom()?;
        let header_footers = self.serialize_header_footers()?;
        // 1) We'll write DocumentContainer at stream offset 0
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        // Allocate a persist ID for the Document itself and set its offset to 0
        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster (top-level record written after Document)
        let master_persist_id = persist_builder.allocate_id();
        let vba_persist_id = self
            .vba_project
            .as_ref()
            .map(|_| persist_builder.allocate_id());

        // 2) Build DocumentContainer
        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        // 2.1) DocumentAtom
        let doc_atom = create_document_atom_with_font_embedding(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
            font_plan.save_with_fonts,
        )?;
        doc_container.write_child(&doc_atom);

        // 2.2) Environment (with FontCollection)
        doc_container.write_child(&font_plan.environment);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // Calculate per-slide shape counts (group + background + user shapes)
        let master_shapes = 6u32;
        let slide_shape_counts: Vec<u32> = self
            .slides
            .iter()
            .map(|s| s.escher_shape_count()) // 2 for group+background, plus user shapes and tables
            .collect();
        // Build DggContainer with BStore if pictures are present
        let dgg = if !self.blip_store.is_empty() {
            let bstore = self.blip_store.store().map_err(WriteError::Io)?;
            super::super::escher::create_dgg_container_with_blips(
                master_shapes,
                &slide_shape_counts,
                &bstore,
            )?
        } else {
            create_dgg_container(master_shapes, &slide_shape_counts)?
        };
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1) referencing MainMaster
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // 2.4) DocInfo List (0x07D0) before SlideListWithText (slides), per POI empty_textbox.ppt
        let docinfo = self.build_docinfo_list(
            vba_persist_id,
            &font_plan.powerpoint10_records,
            modify_password_atom.as_deref(),
        )?;
        doc_container.write_child(&docinfo);

        if let Some(value) = &header_footers.presentation_slides {
            doc_container.write_child(value);
        }
        if let Some(value) = &header_footers.notes_and_handouts {
            doc_container.write_child(value);
        }

        // 2.5) SlideListWithText (SLIDES) referencing each slide by (persist id ref, slide identifier)
        let mut slide_persist_ids = Vec::with_capacity(self.slides.len());
        let mut slwt_entries = Vec::with_capacity(self.slides.len());
        for (i, _slide) in self.slides.iter().enumerate() {
            let pid = persist_builder.allocate_id();
            slide_persist_ids.push(pid);
            let slide_identifier = 256u32 + (i as u32);
            slwt_entries.push((pid, slide_identifier));
        }
        if !slwt_entries.is_empty() {
            use super::super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
        }

        // 2.5.1) Pre-allocate notes persist IDs and build SlideListWithText for notes
        // Per POI: Notes' SlidePersistAtom.slideIdentifier must match Slide's slideIdentifier
        // This is how POI matches notes to slides in findNotesSlides/findSlides
        let mut notes_persist_ids: Vec<Option<u32>> = vec![None; self.slides.len()];
        let mut notes_slwt_entries = Vec::new();
        for (i, slide) in self.slides.iter().enumerate() {
            let has_notes =
                slide.notes.as_ref().is_some_and(|n| !n.is_empty()) || slide.notes_page.is_some();
            if has_notes {
                let notes_pid = persist_builder.allocate_id();
                notes_persist_ids[i] = Some(notes_pid);
                // Use SAME slideIdentifier as the slide (256 + i) for matching!
                let slide_identifier = 256u32 + (i as u32);
                notes_slwt_entries.push((notes_pid, slide_identifier));
            }
        }
        if !notes_slwt_entries.is_empty() {
            use super::super::records::create_slide_list_with_text_notes;
            let slwt_notes = create_slide_list_with_text_notes(&notes_slwt_entries)?;
            doc_container.write_child(&slwt_notes);
        }

        // 2.5.2) ExObjList for hyperlinks and embedded chart objects (if any)
        let chart_plans = self.plan_charts(&mut persist_builder)?;
        let ex_obj_list = super::super::chart::build_ex_obj_list(&self.hyperlinks, &chart_plans)?;
        if !ex_obj_list.is_empty() {
            doc_container.write_child(&ex_obj_list);
        }

        // 2.5.3) SoundCollection for animation, shape, and text-action references.
        let (sound_collection, sound_id_mapping) =
            build_writer_sound_collection(&self.slides, &self.sound_resources)?;
        if !sound_collection.is_empty() {
            doc_container.write_child(&sound_collection);
        }

        // 2.5.4) NamedShows (custom slide shows) in Document container
        if !self.custom_shows.is_empty() {
            let named_shows = super::super::custom_shows::build_named_shows(&self.custom_shows)?;
            if !named_shows.is_empty() {
                doc_container.write_child(&named_shows);
            }
        }

        // 2.6) EndDocument
        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Finalize DocumentContainer and write to stream (offset 0)
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // 3) MainMaster then Slides (top-level after DocumentContainer)
        // 3.1) Write MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let mut master_container = create_main_master_container(&master_ppdrawing)?;
        if let Some(value) = &header_footers.main_master {
            append_child_to_built_container(&mut master_container, value)?;
        }
        let master_offset = ppt_stream.len() as u32;
        persist_builder.set_offset(master_persist_id, master_offset);
        ppt_stream.extend_from_slice(&master_container);

        // 3.2) Slides
        for (i, slide) in self.slides.iter().enumerate() {
            // drawing_id for slides starts from 2 (1 is used by MainMaster)
            let drawing_id = (i as u32) + 2;
            let slide_identifier = 256u32 + (i as u32);

            // Build Slide container with SlideAtom
            let mut slide_container = RecordBuilder::new(0x0F, 0, record_type::SLIDE);
            // SlideAtom (MS-PPT 2.4.7)
            let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
            let mut atom_data = Vec::with_capacity(24);
            // SSlideLayoutAtom: geometry + placeholder types
            atom_data.extend_from_slice(&(SlideLayoutType::Blank as u32).to_le_bytes());
            atom_data.extend_from_slice(&[0u8; 8]); // rgPlaceholderTypes
            // masterIdRef (0x80000000 = reference to master)
            atom_data.extend_from_slice(&0x8000_0000u32.to_le_bytes());
            // notesIdRef: Per POI, this equals NotesAtom.slideID = slideIdentifier
            // Set to the slide's own identifier if notes exist, 0 otherwise
            let notes_id_ref = if notes_persist_ids[i].is_some() {
                slide_identifier // Same value as NotesAtom.slideID
            } else {
                0
            };
            atom_data.extend_from_slice(&notes_id_ref.to_le_bytes());
            // slideFlags: follow master objects/scheme/background
            atom_data.extend_from_slice(&slide_flags::DEFAULT.to_le_bytes());
            atom_data.extend_from_slice(&0u16.to_le_bytes()); // reserved
            slide_atom.write_data(&atom_data);
            slide_container.write_child(&slide_atom.build()?);

            // PPDrawing with Escher DgContainer (including user shapes)
            let escher_shapes: Vec<UserShapeData> = slide
                .shapes
                .iter()
                .map(|s| {
                    convert_shape_to_escher_with_sound_mapping(
                        s,
                        &self.hyperlinks,
                        &sound_id_mapping,
                    )
                })
                .collect();
            // Chart frames referencing this slide's embedded chart objects
            let chart_frames: Vec<super::super::chart::ChartFrame> = chart_plans
                .iter()
                .filter(|plan| plan.slide == i)
                .map(|plan| {
                    let chart = &slide.charts[plan.chart];
                    super::super::chart::ChartFrame::new(
                        chart.x,
                        chart.y,
                        chart.width,
                        chart.height,
                        plan.ex_obj_id,
                    )
                })
                .collect::<Result<Vec<_>, _>>()?;
            let dg = create_dg_container_with_charts(
                drawing_id,
                &escher_shapes,
                &slide.tables,
                &chart_frames,
            )?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // SSSlideInfoAtom for per-slide timing (if set and no transition handles it)
            if let Some(ref timing) = slide.timing {
                let timing_record = super::super::slide_timing::build_slide_timing(timing)?;
                slide_container.write_child(&timing_record);
            }

            if let Some(value) = &header_footers.slides[i] {
                slide_container.write_child(value);
            }

            // ProgTags with PPT10 binary tag (PowerPoint 2002+ features)
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Tag10::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            // BinaryTagData: slide defaults + comments
            let comment_bytes = super::super::comments::build_slide_comments(&slide.comments)?;
            let mut tag_data = BinaryTagData::SLIDE.to_bytes().to_vec();
            tag_data.extend_from_slice(&comment_bytes);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&tag_data);
            prog_bin.write_child(&bin.build()?);
            prog_tags.write_child(&prog_bin.build()?);
            slide_container.write_child(&prog_tags.build()?);

            // Compute this slide's offset in the stream: current top-level length
            let slide_offset = ppt_stream.len() as u32;

            // Track persist pointer (allocate new persist id per slide)
            let persist_id = slide_persist_ids[i];
            persist_builder.set_offset(persist_id, slide_offset);

            // Append slide as top-level record
            let slide_bytes = slide_container.build()?;
            ppt_stream.extend_from_slice(&slide_bytes);
        }

        // 3.3) Notes containers for slides with notes
        for (i, slide) in self.slides.iter().enumerate() {
            if let Some(notes_pid) = notes_persist_ids[i] {
                let notes_offset = ppt_stream.len() as u32;
                persist_builder.set_offset(notes_pid, notes_offset);

                // Per POI: NotesAtom.slideID = slideIdentifier (same as slide's identifier)
                // This equals SlideAtom.notesID and Notes' SlidePersistAtom.slideIdentifier
                let slide_identifier = 256u32 + (i as u32);
                let notes_page = if let Some(page) = &slide.notes_page {
                    let mut page = page.clone();
                    page.slide_id_ref = slide_identifier;
                    page
                } else if let Some(text) = &slide.notes {
                    NotesPage::simple(slide_identifier, text)
                } else {
                    continue;
                };

                // Build notes container (drawing_id continues after slides)
                let notes_drawing_id = (self.slides.len() as u32) + 2 + (i as u32);
                let notes_builder = NotesContainerBuilder::new(notes_page, notes_drawing_id);
                let notes_bytes = notes_builder.build().map_err(std::io::Error::other)?;
                ppt_stream.extend_from_slice(&notes_bytes);
            }
        }

        // 3.4) ExOleObjStg persisted storages for embedded chart objects
        for plan in &chart_plans {
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(plan.persist_id, offset);
            let storage = super::super::chart::chart_storage_record(
                &self.slides[plan.slide].charts[plan.chart].workbook,
            )?;
            ppt_stream.extend_from_slice(&storage);
        }

        if let (Some(persist_id), Some(storage)) = (vba_persist_id, &self.vba_project) {
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            let record = storage
                .to_record_bytes()
                .map_err(|error| WriteError::InvalidData(error.to_string()))?;
            ppt_stream.extend_from_slice(&record);
        }

        let pictures_stream = if self.blip_store.is_empty() {
            None
        } else {
            Some(self.blip_store.delay().map_err(WriteError::Io)?)
        };
        #[cfg(feature = "encryption")]
        let mut pictures_stream = pictures_stream;
        #[cfg(feature = "encryption")]
        let encryption = self.prepare_encryption()?;
        #[cfg(feature = "encryption")]
        let encryption_session_id = if let Some(encryption) = &encryption {
            let persist_id = persist_builder.allocate_id();
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            ppt_stream.extend_from_slice(&encryption.session_record);
            Some(persist_id)
        } else {
            None
        };

        // 4) PersistPtrIncrementalBlock (6002) then single UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        #[cfg(feature = "encryption")]
        let mut user_edit = user_edit;
        #[cfg(feature = "encryption")]
        if let Some(session_id) = encryption_session_id {
            user_edit = user_edit.with_encryption_session(session_id);
        }
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        // 5) Build Current User and property streams
        #[cfg(feature = "encryption")]
        let encrypted = encryption.is_some();
        #[cfg(not(feature = "encryption"))]
        let encrypted = false;
        let current_user = build_current_user_stream(user_edit_offset, encrypted);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        #[cfg(feature = "encryption")]
        if let (Some(encryption), Some(session_id)) = (&encryption, encryption_session_id) {
            encrypt_powerpoint_document_for_write(
                &mut ppt_stream,
                persist_dir_offset as usize,
                user_edit_offset as usize,
                session_id,
                &encryption.crypto,
            )
            .map_err(WriteError::InvalidData)?;
            if let Some(pictures) = &mut pictures_stream {
                encrypt_pictures_for_write(pictures, &encryption.crypto)
                    .map_err(WriteError::InvalidData)?;
            }
        }

        // 6) Write OLE streams
        let mut ole_writer = OleWriter::new();
        // Set root CLSID to PowerPoint V8
        ole_writer.set_root_clsid([
            0x10, 0x8D, 0x81, 0x64, 0x9B, 0x4F, 0xCF, 0x11, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9,
            0x29, 0xE8,
        ]);
        ole_writer.create_stream(&["PowerPoint Document"], &ppt_stream)?;
        ole_writer.create_stream(&["Current User"], &current_user)?;
        ole_writer.create_stream(&["\u{0005}SummaryInformation"], &summary_info)?;
        ole_writer.create_stream(&["\u{0005}DocumentSummaryInformation"], &doc_summary)?;

        // Pictures stream (per POI: separate stream for BLIP data)
        if let Some(pictures_stream) = &pictures_stream {
            ole_writer.create_stream(&["Pictures"], pictures_stream)?;
        }

        ole_writer.save(path)?;

        Ok(())
    }

    /// Write presentation to an in-memory buffer
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer (must support Write + Seek)
    ///
    /// # Returns
    ///
    /// * `Result<(), WriteError>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), WriteError> {
        self.validate_encryption()?;
        self.validate_references()?;
        let font_plan = self.font_publication_plan()?;
        let modify_password_atom = self.build_modify_password_atom()?;
        let header_footers = self.serialize_header_footers()?;
        // Same logic as save(), but writing to provided writer
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster
        let master_persist_id = persist_builder.allocate_id();
        let vba_persist_id = self
            .vba_project
            .as_ref()
            .map(|_| persist_builder.allocate_id());

        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        let doc_atom = create_document_atom_with_font_embedding(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
            font_plan.save_with_fonts,
        )?;
        doc_container.write_child(&doc_atom);
        // 2.2) Environment (with FontCollection)
        doc_container.write_child(&font_plan.environment);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // Calculate per-slide shape counts (group + background + user shapes)
        let master_shapes = 6u32;
        let slide_shape_counts: Vec<u32> =
            self.slides.iter().map(|s| s.escher_shape_count()).collect();
        // Build DggContainer with BStore if pictures are present
        let dgg = if !self.blip_store.is_empty() {
            let bstore = self.blip_store.store().map_err(WriteError::Io)?;
            super::super::escher::create_dgg_container_with_blips(
                master_shapes,
                &slide_shape_counts,
                &bstore,
            )?
        } else {
            create_dgg_container(master_shapes, &slide_shape_counts)?
        };
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1)
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // DocInfo List before SlideListWithText (slides), matching POI empty_textbox.ppt
        let docinfo = self.build_docinfo_list(
            vba_persist_id,
            &font_plan.powerpoint10_records,
            modify_password_atom.as_deref(),
        )?;
        doc_container.write_child(&docinfo);

        if let Some(value) = &header_footers.presentation_slides {
            doc_container.write_child(value);
        }
        if let Some(value) = &header_footers.notes_and_handouts {
            doc_container.write_child(value);
        }

        // SlideListWithText (SLIDES) for non-empty presentations
        let mut slide_persist_ids = Vec::with_capacity(self.slides.len());
        let mut slwt_entries = Vec::with_capacity(self.slides.len());
        for (i, _slide) in self.slides.iter().enumerate() {
            let pid = persist_builder.allocate_id();
            slide_persist_ids.push(pid);
            let slide_identifier = 256u32 + (i as u32);
            slwt_entries.push((pid, slide_identifier));
        }
        if !slwt_entries.is_empty() {
            use super::super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
        }

        // ExObjList for hyperlinks and embedded chart objects (if any)
        let chart_plans = self.plan_charts(&mut persist_builder)?;
        let ex_obj_list = super::super::chart::build_ex_obj_list(&self.hyperlinks, &chart_plans)?;
        if !ex_obj_list.is_empty() {
            doc_container.write_child(&ex_obj_list);
        }

        // SoundCollection for animation, shape, and text-action references.
        let (sound_collection, sound_id_mapping) =
            build_writer_sound_collection(&self.slides, &self.sound_resources)?;
        if !sound_collection.is_empty() {
            doc_container.write_child(&sound_collection);
        }

        // NamedShows (custom slide shows) in Document container
        if !self.custom_shows.is_empty() {
            let named_shows = super::super::custom_shows::build_named_shows(&self.custom_shows)?;
            if !named_shows.is_empty() {
                doc_container.write_child(&named_shows);
            }
        }

        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Write finalized DocumentContainer
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // Then write MainMaster and slides as top-level records
        // MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let mut master_container = create_main_master_container(&master_ppdrawing)?;
        if let Some(value) = &header_footers.main_master {
            append_child_to_built_container(&mut master_container, value)?;
        }
        let master_offset = ppt_stream.len() as u32;
        persist_builder.set_offset(master_persist_id, master_offset);
        ppt_stream.extend_from_slice(&master_container);

        // Slides
        for (i, slide) in self.slides.iter().enumerate() {
            let drawing_id = (i as u32) + 2; // 1 reserved for master

            let mut slide_container = RecordBuilder::new(0x0F, 0, record_type::SLIDE);
            // SlideAtom (MS-PPT 2.4.7)
            let mut slide_atom = RecordBuilder::new(0x02, 0, record_type::SLIDE_ATOM);
            let mut atom_data = Vec::with_capacity(24);
            atom_data.extend_from_slice(&(SlideLayoutType::Blank as u32).to_le_bytes());
            atom_data.extend_from_slice(&[0u8; 8]); // rgPlaceholderTypes
            atom_data.extend_from_slice(&0x8000_0000u32.to_le_bytes()); // masterIdRef
            atom_data.extend_from_slice(&0u32.to_le_bytes()); // notesIdRef
            atom_data.extend_from_slice(&slide_flags::DEFAULT.to_le_bytes());
            atom_data.extend_from_slice(&0u16.to_le_bytes()); // reserved
            slide_atom.write_data(&atom_data);
            slide_container.write_child(&slide_atom.build()?);

            // PPDrawing with Escher DgContainer (including user shapes)
            let escher_shapes: Vec<UserShapeData> = slide
                .shapes
                .iter()
                .map(|s| {
                    convert_shape_to_escher_with_sound_mapping(
                        s,
                        &self.hyperlinks,
                        &sound_id_mapping,
                    )
                })
                .collect();
            // Chart frames referencing this slide's embedded chart objects
            let chart_frames: Vec<super::super::chart::ChartFrame> = chart_plans
                .iter()
                .filter(|plan| plan.slide == i)
                .map(|plan| {
                    let chart = &slide.charts[plan.chart];
                    super::super::chart::ChartFrame::new(
                        chart.x,
                        chart.y,
                        chart.width,
                        chart.height,
                        plan.ex_obj_id,
                    )
                })
                .collect::<Result<Vec<_>, _>>()?;
            let dg = create_dg_container_with_charts(
                drawing_id,
                &escher_shapes,
                &slide.tables,
                &chart_frames,
            )?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // SSSlideInfoAtom for per-slide timing (if set)
            if let Some(ref timing) = slide.timing {
                let timing_record = super::super::slide_timing::build_slide_timing(timing)?;
                slide_container.write_child(&timing_record);
            }

            if let Some(value) = &header_footers.slides[i] {
                slide_container.write_child(value);
            }

            // ProgTags with PPT10 binary tag
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Tag10::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            // BinaryTagData: slide defaults + comments
            let comment_bytes = super::super::comments::build_slide_comments(&slide.comments)?;
            let mut tag_data = BinaryTagData::SLIDE.to_bytes().to_vec();
            tag_data.extend_from_slice(&comment_bytes);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&tag_data);
            prog_bin.write_child(&bin.build()?);
            prog_tags.write_child(&prog_bin.build()?);
            slide_container.write_child(&prog_tags.build()?);

            let slide_offset = ppt_stream.len() as u32;
            let persist_id = slide_persist_ids[i];
            persist_builder.set_offset(persist_id, slide_offset);

            let slide_bytes = slide_container.build()?;
            ppt_stream.extend_from_slice(&slide_bytes);
        }

        // 3.3) Notes containers - DISABLED for testing
        // Notes need more work - SlideListWithText instance=2, proper linking

        // 3.4) ExOleObjStg persisted storages for embedded chart objects
        for plan in &chart_plans {
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(plan.persist_id, offset);
            let storage = super::super::chart::chart_storage_record(
                &self.slides[plan.slide].charts[plan.chart].workbook,
            )?;
            ppt_stream.extend_from_slice(&storage);
        }

        if let (Some(persist_id), Some(storage)) = (vba_persist_id, &self.vba_project) {
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            let record = storage
                .to_record_bytes()
                .map_err(|error| WriteError::InvalidData(error.to_string()))?;
            ppt_stream.extend_from_slice(&record);
        }

        let pictures_stream = if self.blip_store.is_empty() {
            None
        } else {
            Some(self.blip_store.delay().map_err(WriteError::Io)?)
        };
        #[cfg(feature = "encryption")]
        let mut pictures_stream = pictures_stream;
        #[cfg(feature = "encryption")]
        let encryption = self.prepare_encryption()?;
        #[cfg(feature = "encryption")]
        let encryption_session_id = if let Some(encryption) = &encryption {
            let persist_id = persist_builder.allocate_id();
            let offset = u32::try_from(ppt_stream.len()).map_err(|_| {
                WriteError::InvalidData("PPT document stream exceeds 4 GiB".to_string())
            })?;
            persist_builder.set_offset(persist_id, offset);
            ppt_stream.extend_from_slice(&encryption.session_record);
            Some(persist_id)
        } else {
            None
        };

        // PersistPtrHolder and UserEditAtom
        let persist_dir_offset = ppt_stream.len() as u32;
        let persist_dir_block = persist_builder.generate_record();
        ppt_stream.extend_from_slice(&persist_dir_block);

        let user_edit = UserEditAtom::new_minimal(
            persist_dir_offset,
            doc_persist_id,
            persist_builder.persist_id_seed(),
            self.slides.len() as u32,
        );
        #[cfg(feature = "encryption")]
        let mut user_edit = user_edit;
        #[cfg(feature = "encryption")]
        if let Some(session_id) = encryption_session_id {
            user_edit = user_edit.with_encryption_session(session_id);
        }
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        #[cfg(feature = "encryption")]
        let encrypted = encryption.is_some();
        #[cfg(not(feature = "encryption"))]
        let encrypted = false;
        let current_user = build_current_user_stream(user_edit_offset, encrypted);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        #[cfg(feature = "encryption")]
        if let (Some(encryption), Some(session_id)) = (&encryption, encryption_session_id) {
            encrypt_powerpoint_document_for_write(
                &mut ppt_stream,
                persist_dir_offset as usize,
                user_edit_offset as usize,
                session_id,
                &encryption.crypto,
            )
            .map_err(WriteError::InvalidData)?;
            if let Some(pictures) = &mut pictures_stream {
                encrypt_pictures_for_write(pictures, &encryption.crypto)
                    .map_err(WriteError::InvalidData)?;
            }
        }

        let mut ole_writer = OleWriter::new();
        ole_writer.set_root_clsid([
            0x10, 0x8D, 0x81, 0x64, 0x9B, 0x4F, 0xCF, 0x11, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9,
            0x29, 0xE8,
        ]);
        ole_writer.create_stream(&["PowerPoint Document"], &ppt_stream)?;
        ole_writer.create_stream(&["Current User"], &current_user)?;
        ole_writer.create_stream(&["\u{0005}SummaryInformation"], &summary_info)?;
        ole_writer.create_stream(&["\u{0005}DocumentSummaryInformation"], &doc_summary)?;

        // Pictures stream (per POI: separate stream for BLIP data)
        if let Some(pictures_stream) = &pictures_stream {
            ole_writer.create_stream(&["Pictures"], pictures_stream)?;
        }

        ole_writer.write_to(writer)?;

        Ok(())
    }

    // Helper methods for PPT writer:
    // The following are implemented via the modular components:
    // - Generating PPT record headers and containers
    // - Building Escher drawing records (DggContainer, DgContainer, etc.)
    // - Creating shape records (ClientData, ClientAnchor, etc.)
    // - Building text run records (TextCharsAtom, TextBytesAtom)
    // - Generating PersistPtr directory
    // - Creating CurrentUser stream
    // - Building SlideAtom and NotesAtom structures
    // - Managing master slides and layouts
    //
    // For production use, the PPTX writer is fully implemented and recommended.
}
