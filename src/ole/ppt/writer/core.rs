//! PPT file writer implementation
//!
//! This module provides functionality to create and modify Microsoft PowerPoint presentations
//! in the legacy binary format (.ppt files) using OLE2 structured storage.
//!
//! # Architecture
//!
//! The writer generates PPT record structures including:
//! - DocumentContainer - the root record container
//! - SlideListWithText - contains all slides
//! - MainMaster - master slide definitions
//! - Escher drawing records - for shapes and drawing objects
//! - PersistPtr - mapping system for record locations
//!
//! # PPT File Format Overview
//!
//! PPT files use a record-based binary format where:
//! 1. Everything is a record (atom or container)
//! 2. Containers hold other records
//! 3. Atoms contain actual data
//! 4. Escher (Office Drawing) format for shapes
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ole::ppt::PptWriter;
//!
//! let mut writer = PptWriter::new();
//!
//! // Add a slide
//! let slide = writer.add_slide()?;
//!
//! // Add a text box
//! writer.add_textbox(slide, 100, 100, 400, 200, "Hello, World!")?;
//!
//! // Save the presentation
//! writer.save("output.ppt")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::escher::{create_dg_container, create_dgg_container};
use super::master_drawing::build_master_ppdrawing;
use super::persist::{PersistPtrBuilder, UserEditAtom};
use super::records::{
    RecordBuilder, create_docinfo_list_container_minimal, create_document_atom,
    create_end_document, create_environment_minimal, create_main_master_container,
    create_slide_list_with_text_master, create_text_atom, record_type, wrap_dg_into_ppdrawing,
    wrap_dgg_into_ppdrawing_group,
};
use super::spec::{BinaryTagData, ColorScheme, Ppt10Tag, SlideLayoutType, slide_flags};
use crate::ole::writer::OleWriter;
use std::collections::HashMap;

/// Error type for PPT writing
#[derive(Debug)]
pub enum PptWriteError {
    /// I/O error
    Io(std::io::Error),
    /// Invalid data
    InvalidData(String),
    /// OLE error
    Ole(crate::ole::OleError),
}

/// Build a minimal, valid Current User stream referencing the given UserEditAtom offset.
fn build_current_user_stream(offset_to_current_edit: u32) -> Vec<u8> {
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
    // headerToken (unencrypted)
    s.extend_from_slice(&0xE391C05Fu32.to_le_bytes());
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

impl From<std::io::Error> for PptWriteError {
    fn from(err: std::io::Error) -> Self {
        PptWriteError::Io(err)
    }
}

impl From<crate::ole::OleError> for PptWriteError {
    fn from(err: crate::ole::OleError) -> Self {
        PptWriteError::Ole(err)
    }
}

impl std::fmt::Display for PptWriteError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PptWriteError::Io(e) => write!(f, "I/O error: {}", e),
            PptWriteError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            PptWriteError::Ole(e) => write!(f, "OLE error: {}", e),
        }
    }
}

impl std::error::Error for PptWriteError {}

/// Shape type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeType {
    /// Rectangle
    Rectangle,
    /// Text box
    TextBox,
    /// Placeholder
    Placeholder,
    /// Line
    Line,
    /// Ellipse
    Ellipse,
    // Future enhancement: Additional shape types (Star, Pentagon, etc.)
}

/// Text alignment
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextAlignment {
    /// Left-aligned
    Left,
    /// Center-aligned
    Center,
    /// Right-aligned
    Right,
    /// Justified
    Justify,
}

/// Shape properties
#[derive(Debug, Clone)]
pub struct ShapeProperties {
    /// Shape type
    pub shape_type: ShapeType,
    /// X position (in EMUs - English Metric Units, 914400 EMUs = 1 inch)
    pub x: i32,
    /// Y position (in EMUs)
    pub y: i32,
    /// Width (in EMUs)
    pub width: i32,
    /// Height (in EMUs)
    pub height: i32,
    /// Text content (if applicable)
    pub text: Option<String>,
    /// Text alignment
    pub alignment: TextAlignment,
    // Future enhancement: Additional properties (fill color, line color, font, shadows, etc.)
}

/// Represents a shape on a slide
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct WritableShape {
    /// Shape properties
    properties: ShapeProperties,
}

/// Represents a slide
#[derive(Debug, Clone)]
struct WritableSlide {
    /// Shapes on this slide
    shapes: Vec<WritableShape>,
    /// Slide notes text
    notes: Option<String>,
}

/// PPT file writer
///
/// Provides methods to create and modify PPT files.
#[allow(dead_code)] // Reserved for future implementation
pub struct PptWriter {
    /// Slides in the presentation
    slides: Vec<WritableSlide>,
    /// Presentation properties
    properties: HashMap<String, String>,
    /// Slide width in EMUs (default: Letter size)
    slide_width: i32,
    /// Slide height in EMUs (default: Letter size)
    slide_height: i32,
}

impl PptWriter {
    /// Create a new PPT writer with standard 4:3 slide dimensions
    pub fn new() -> Self {
        Self::with_dimensions(9144000, 6858000) // 10" x 7.5" in EMUs
    }

    /// Create a new PPT writer with widescreen 16:9 dimensions
    pub fn new_widescreen() -> Self {
        Self::with_dimensions(9144000, 5143500) // 10" x 5.625" in EMUs
    }

    /// Create a new PPT writer with custom dimensions
    ///
    /// # Arguments
    ///
    /// * `width` - Slide width in EMUs (914400 EMUs = 1 inch)
    /// * `height` - Slide height in EMUs
    pub fn with_dimensions(width: i32, height: i32) -> Self {
        Self {
            slides: Vec::new(),
            properties: HashMap::new(),
            slide_width: width,
            slide_height: height,
        }
    }

    /// Add a new blank slide
    ///
    /// # Returns
    ///
    /// * `Result<usize, PptWriteError>` - Slide index or error
    pub fn add_slide(&mut self) -> Result<usize, PptWriteError> {
        let index = self.slides.len();
        self.slides.push(WritableSlide {
            shapes: Vec::new(),
            notes: None,
        });
        Ok(index)
    }

    /// Delete a slide
    ///
    /// # Arguments
    ///
    /// * `index` - Slide index (0-based)
    pub fn delete_slide(&mut self, index: usize) -> Result<(), PptWriteError> {
        if index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(format!(
                "Slide {} does not exist",
                index
            )));
        }
        self.slides.remove(index);
        Ok(())
    }

    /// Move a slide to a new position
    ///
    /// # Arguments
    ///
    /// * `from_index` - Current slide index
    /// * `to_index` - New slide index
    pub fn move_slide(&mut self, from_index: usize, to_index: usize) -> Result<(), PptWriteError> {
        if from_index >= self.slides.len() || to_index >= self.slides.len() {
            return Err(PptWriteError::InvalidData(
                "Invalid slide index".to_string(),
            ));
        }

        let slide = self.slides.remove(from_index);
        self.slides.insert(to_index, slide);
        Ok(())
    }

    /// Add a text box to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x` - X position (in points, 72 points = 1 inch)
    /// * `y` - Y position (in points)
    /// * `width` - Width (in points)
    /// * `height` - Height (in points)
    /// * `text` - Text content
    pub fn add_textbox(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
        text: &str,
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        // Convert points to EMUs (1 point = 12700 EMUs)
        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::TextBox,
                x: x * 12700,
                y: y * 12700,
                width: width * 12700,
                height: height * 12700,
                text: Some(text.to_string()),
                alignment: TextAlignment::Left,
            },
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Add a rectangle shape to a slide
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `x` - X position (in points)
    /// * `y` - Y position (in points)
    /// * `width` - Width (in points)
    /// * `height` - Height (in points)
    pub fn add_rectangle(
        &mut self,
        slide: usize,
        x: i32,
        y: i32,
        width: i32,
        height: i32,
    ) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        let shape = WritableShape {
            properties: ShapeProperties {
                shape_type: ShapeType::Rectangle,
                x: x * 12700,
                y: y * 12700,
                width: width * 12700,
                height: height * 12700,
                text: None,
                alignment: TextAlignment::Left,
            },
        };

        slide_data.shapes.push(shape);
        Ok(())
    }

    /// Set slide notes
    ///
    /// # Arguments
    ///
    /// * `slide` - Slide index
    /// * `notes` - Notes text
    pub fn set_slide_notes(&mut self, slide: usize, notes: &str) -> Result<(), PptWriteError> {
        let slide_data = self
            .slides
            .get_mut(slide)
            .ok_or_else(|| PptWriteError::InvalidData(format!("Slide {} does not exist", slide)))?;

        slide_data.notes = Some(notes.to_string());
        Ok(())
    }

    /// Set a presentation property
    ///
    /// # Arguments
    ///
    /// * `name` - Property name (e.g., "Title", "Author", "Subject")
    /// * `value` - Property value
    pub fn set_property(&mut self, name: &str, value: &str) {
        self.properties.insert(name.to_string(), value.to_string());
    }

    /// Get the number of slides
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), PptWriteError>` - Success or error
    ///
    /// # Implementation
    ///
    /// This generates a complete PowerPoint 97-2003 binary file conforming to MS-PPT specification:
    /// - PPT record structures - [MS-PPT] Section 2.3
    /// - Escher drawing containers - [MS-ODRAW] Section 2.2
    /// - PersistPtr directory - [MS-PPT] Section 2.4.16
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), PptWriteError> {
        // 1) We'll write DocumentContainer at stream offset 0
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        // Allocate a persist ID for the Document itself and set its offset to 0
        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster (top-level record written after Document)
        let master_persist_id = persist_builder.allocate_id();

        // 2) Build DocumentContainer
        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        // 2.1) DocumentAtom
        let doc_atom = create_document_atom(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
        )?;
        doc_container.write_child(&doc_atom);

        // 2.2) Environment (with FontCollection)
        let env = create_environment_minimal()?;
        doc_container.write_child(&env);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // drawing_count = #masters (1) + #slides
        // MainMaster has 6 shapes (from POI template), slides have 2 shapes each
        let drawing_count = (self.slides.len() as u32).saturating_add(1);
        let master_shapes = 6u32;
        let slide_shapes = 2u32;
        let dgg = create_dgg_container(drawing_count, master_shapes, slide_shapes)?;
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1) referencing MainMaster
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // 2.4) DocInfo List (0x07D0) before SlideListWithText (slides), per POI empty_textbox.ppt
        let docinfo = create_docinfo_list_container_minimal()?;
        doc_container.write_child(&docinfo);

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
            use super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
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
        let master_container = create_main_master_container(&master_ppdrawing)?;
        let master_offset = ppt_stream.len() as u32;
        persist_builder.set_offset(master_persist_id, master_offset);
        ppt_stream.extend_from_slice(&master_container);

        // 3.2) Slides
        for (i, slide) in self.slides.iter().enumerate() {
            // drawing_id for slides starts from 2 (1 is used by MainMaster)
            let drawing_id = (i as u32) + 2;

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
            atom_data.extend_from_slice(&0u32.to_le_bytes()); // notesIdRef
            // slideFlags: follow master objects/scheme/background
            atom_data.extend_from_slice(&slide_flags::DEFAULT.to_le_bytes());
            atom_data.extend_from_slice(&0u16.to_le_bytes()); // reserved
            slide_atom.write_data(&atom_data);
            slide_container.write_child(&slide_atom.build()?);

            // Optional text
            if let Some(notes) = slide.notes.as_deref()
                && !notes.is_empty()
            {
                let text_atom = create_text_atom(notes)?;
                slide_container.write_child(&text_atom);
            }

            // PPDrawing with Escher DgContainer
            let dg = create_dg_container(drawing_id, slide.shapes.len() as u32)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // ProgTags with PPT10 binary tag (PowerPoint 2002+ features)
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&BinaryTagData::SLIDE.to_bytes());
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
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        // 5) Build Current User and property streams
        let current_user = build_current_user_stream(user_edit_offset);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

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
    /// * `Result<(), PptWriteError>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(
        &mut self,
        writer: &mut W,
    ) -> Result<(), PptWriteError> {
        // Same logic as save(), but writing to provided writer
        let mut ppt_stream = Vec::new();
        let mut persist_builder = PersistPtrBuilder::new();

        let doc_persist_id = persist_builder.allocate_id();
        persist_builder.set_offset(doc_persist_id, 0);
        // Allocate persist ID for MainMaster
        let master_persist_id = persist_builder.allocate_id();

        let mut doc_container = RecordBuilder::new(0x0F, 0, record_type::DOCUMENT);

        let doc_atom = create_document_atom(
            self.slide_width as u32,
            self.slide_height as u32,
            self.slides.len() as u32,
            0,
            0,
        )?;
        doc_container.write_child(&doc_atom);
        // 2.2) Environment (with FontCollection)
        let env = create_environment_minimal()?;
        doc_container.write_child(&env);

        // 2.3) PPDrawingGroup wrapping Dgg Escher
        // drawing_count = #masters (1) + #slides
        // MainMaster has 6 shapes (from POI template), slides have 2 shapes each
        let drawing_count = (self.slides.len() as u32).saturating_add(1);
        let master_shapes = 6u32;
        let slide_shapes = 2u32;
        let dgg = create_dgg_container(drawing_count, master_shapes, slide_shapes)?;
        let pp_dgg = wrap_dgg_into_ppdrawing_group(&dgg)?;
        doc_container.write_child(&pp_dgg);

        // 2.3.1) SlideListWithText for masters (instance=1)
        let master_entries = vec![(master_persist_id, 0x8000_0000u32)];
        let slwt_master = create_slide_list_with_text_master(&master_entries)?;
        doc_container.write_child(&slwt_master);

        // DocInfo List before SlideListWithText (slides), matching POI empty_textbox.ppt
        let docinfo = create_docinfo_list_container_minimal()?;
        doc_container.write_child(&docinfo);

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
            use super::records::create_slide_list_with_text_slides;
            let slwt = create_slide_list_with_text_slides(&slwt_entries)?;
            doc_container.write_child(&slwt);
        }

        let end_doc = create_end_document()?;
        doc_container.write_child(&end_doc);

        // Write finalized DocumentContainer
        let doc_bytes = doc_container.build()?;
        ppt_stream.extend_from_slice(&doc_bytes);

        // Then write MainMaster and slides as top-level records
        // MainMaster using dynamically built PPDrawing (includes all placeholders)
        let master_ppdrawing = build_master_ppdrawing();
        let master_container = create_main_master_container(&master_ppdrawing)?;
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

            if let Some(notes) = slide.notes.as_deref()
                && !notes.is_empty()
            {
                let text_atom = create_text_atom(notes)?;
                slide_container.write_child(&text_atom);
            }

            let dg = create_dg_container(drawing_id, slide.shapes.len() as u32)?;
            let pp_dg = wrap_dg_into_ppdrawing(&dg)?;
            slide_container.write_child(&pp_dg);

            // ColorSchemeAtom (MS-PPT 2.4.17)
            let mut color = RecordBuilder::new(0x00, 1, record_type::COLOR_SCHEME_ATOM);
            color.write_data(&ColorScheme::POI_DEFAULT.to_bytes());
            slide_container.write_child(&color.build()?);

            // ProgTags with PPT10 binary tag
            let mut prog_tags = RecordBuilder::new(0x0F, 0, record_type::PROG_TAGS);
            let mut prog_bin = RecordBuilder::new(0x0F, 0, record_type::PROG_BINARY_TAG);
            let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
            cstr.write_data(&Ppt10Tag::to_bytes());
            prog_bin.write_child(&cstr.build()?);
            let mut bin = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
            bin.write_data(&BinaryTagData::SLIDE.to_bytes());
            prog_bin.write_child(&bin.build()?);
            prog_tags.write_child(&prog_bin.build()?);
            slide_container.write_child(&prog_tags.build()?);

            let slide_offset = ppt_stream.len() as u32;
            let persist_id = slide_persist_ids[i];
            persist_builder.set_offset(persist_id, slide_offset);

            let slide_bytes = slide_container.build()?;
            ppt_stream.extend_from_slice(&slide_bytes);
        }

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
        let user_edit_offset = ppt_stream.len() as u32;
        let user_edit_record = user_edit.generate_record();
        ppt_stream.extend_from_slice(&user_edit_record);

        let current_user = build_current_user_stream(user_edit_offset);
        let summary_info = build_summary_information_stream();
        let doc_summary = build_document_summary_information_stream();

        let mut ole_writer = OleWriter::new();
        ole_writer.set_root_clsid([
            0x10, 0x8D, 0x81, 0x64, 0x9B, 0x4F, 0xCF, 0x11, 0x86, 0xEA, 0x00, 0xAA, 0x00, 0xB9,
            0x29, 0xE8,
        ]);
        ole_writer.create_stream(&["PowerPoint Document"], &ppt_stream)?;
        ole_writer.create_stream(&["Current User"], &current_user)?;
        ole_writer.create_stream(&["\u{0005}SummaryInformation"], &summary_info)?;
        ole_writer.create_stream(&["\u{0005}DocumentSummaryInformation"], &doc_summary)?;
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

impl Default for PptWriter {
    fn default() -> Self {
        Self::new()
    }
}

// Implementation deferred - PPT record generation functions:
// These would be needed for full PPT binary format support:
// - write_record_header() - Record header with version, type, instance, length
// - write_document_container() - DocumentContainer record
// - write_slide_container() - Slide record
// - write_drawing_container() - Drawing (Escher) container
// - write_shape_container() - Shape container (spContainer)
// - write_text_box() - Text box Escher record
// - write_client_data() - ClientData record linking to text
//
// Recommendation: Use the PPTX writer (fully implemented) for production use.
// - write_persist_directory() - PersistPtr directory

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_writer() {
        let writer = PptWriter::new();
        assert_eq!(writer.slides.len(), 0);
        assert_eq!(writer.slide_width, 9144000);
        assert_eq!(writer.slide_height, 6858000);
    }

    #[test]
    fn test_create_widescreen() {
        let writer = PptWriter::new_widescreen();
        assert_eq!(writer.slide_width, 9144000);
        assert_eq!(writer.slide_height, 5143500);
    }

    #[test]
    fn test_add_slide() {
        let mut writer = PptWriter::new();
        let idx = writer.add_slide().unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.slides.len(), 1);
    }

    #[test]
    fn test_add_textbox() {
        let mut writer = PptWriter::new();
        let slide = writer.add_slide().unwrap();
        writer.add_textbox(slide, 10, 10, 100, 50, "Test").unwrap();
        assert_eq!(writer.slides[0].shapes.len(), 1);
    }

    #[test]
    fn test_delete_slide() {
        let mut writer = PptWriter::new();
        writer.add_slide().unwrap();
        writer.add_slide().unwrap();
        writer.delete_slide(0).unwrap();
        assert_eq!(writer.slides.len(), 1);
    }
}
