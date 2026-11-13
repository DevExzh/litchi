//! Notes slide support for PPT files
//!
//! This module handles speaker notes pages in PowerPoint presentations.
//! Notes pages contain the slide thumbnail and notes text area.
//!
//! Reference: [MS-PPT] Section 2.5.5 - NotesContainer

// =============================================================================
// PPT Record Types for Notes
// =============================================================================

/// Record types for notes
pub mod record_type {
    /// Notes container (`RT_Notes`)
    pub const NOTES: u16 = 0x03F0;
    /// `NotesAtom`
    pub const NOTES_ATOM: u16 = 0x03F1;
    /// `PPDrawing` (contains Escher records)
    pub const PP_DRAWING: u16 = 0x040C;
    /// `TextHeader` (type of text)
    pub const TEXT_HEADER_ATOM: u16 = 0x0F9F;
    /// `TextCharsAtom` (UTF-16 text)
    pub const TEXT_CHARS_ATOM: u16 = 0x0FA0;
    /// `TextBytesAtom` (ASCII text)
    pub const TEXT_BYTES_ATOM: u16 = 0x0FA8;
    /// `StyleTextPropAtom`
    pub const STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;
    /// `ColorSchemeAtom`
    pub const COLOR_SCHEME_ATOM: u16 = 0x07F0;
    /// `SlideListWithText` for notes (instance=2)
    pub const SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
    /// `SlidePersistAtom`
    pub const SLIDE_PERSIST_ATOM: u16 = 0x03F3;
    /// `TextRuler`
    pub const TEXT_RULER_ATOM: u16 = 0x0FA2;
    /// `NotesTextViewInfo`
    pub const NOTES_TEXT_VIEW_INFO: u16 = 0x0413;
}

use std::io::Write;
use zerocopy::IntoBytes as _;
use zerocopy_derive::{Immutable, IntoBytes, KnownLayout};

use litchi_core::unit::ppt_master_i64_to_emu_i32;

use super::text_format::Paragraph;

// =============================================================================
// NotesAtom (MS-PPT 2.5.5.1)
// =============================================================================

/// `NotesAtom` structure (8 bytes)
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesAtom` is the established public API name matching the MS-PPT record name"
)]
#[derive(Debug, Clone, Copy, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct NotesAtom {
    /// Slide persist ID reference
    pub slide_id_ref: u32,
    /// Notes flags
    pub flags: u16,
    /// Reserved
    pub reserved: u16,
}

/// Notes flags
#[allow(
    clippy::module_name_repetitions,
    reason = "`notes_flags` mirrors the flag field of `NotesAtom`; renaming it would break downstream crates"
)]
pub mod notes_flags {
    /// Follow master objects
    pub const FOLLOW_MASTER_OBJECTS: u16 = 0x0001;
    /// Follow master color scheme
    pub const FOLLOW_MASTER_SCHEME: u16 = 0x0002;
    /// Follow master background
    pub const FOLLOW_MASTER_BACKGROUND: u16 = 0x0004;
}

impl NotesAtom {
    /// Size of the atom data
    pub const SIZE: usize = 8;

    /// Create a new `NotesAtom`
    #[must_use]
    pub fn new(slide_id_ref: u32) -> Self {
        Self {
            slide_id_ref,
            flags: 0x0006, // fFollowMasterObjects | fFollowMasterScheme
            reserved: 0,
        }
    }
}

// =============================================================================
// Placeholder Types for Notes
// =============================================================================

/// Placeholder types used in notes pages
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesPlaceholderType` is the established public API name; renaming it would break downstream crates"
)]
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NotesPlaceholderType {
    /// Slide image placeholder
    SlideImage = 0x0B,
    /// Notes body placeholder
    NotesBody = 0x0C,
    /// Header placeholder
    Header = 0x0A,
    /// Footer placeholder
    Footer = 0x09,
    /// Slide number placeholder
    SlideNumber = 0x08,
    /// Date placeholder
    Date = 0x07,
}

// =============================================================================
// Notes Page Definition
// =============================================================================

/// Complete notes page definition
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesPage` is the established public API name; renaming it would break downstream crates"
)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "the four bools map one-to-one to independent header/footer/slide-number/date visibility flags of the notes page format"
)]
#[derive(Debug, Clone)]
pub struct NotesPage {
    /// Persist ID (assigned during save)
    pub persist_id: u32,
    /// Reference to the associated slide's persist ID
    pub slide_id_ref: u32,
    /// Notes text (may contain multiple paragraphs)
    pub text: Vec<Paragraph>,
    /// Whether to show header
    pub show_header: bool,
    /// Whether to show footer
    pub show_footer: bool,
    /// Whether to show slide number
    pub show_slide_number: bool,
    /// Whether to show date
    pub show_date: bool,
    /// Header text
    pub header_text: Option<String>,
    /// Footer text
    pub footer_text: Option<String>,
}

impl NotesPage {
    /// Create a new notes page for a slide
    #[must_use]
    pub fn new(slide_id_ref: u32) -> Self {
        Self {
            persist_id: 0,
            slide_id_ref,
            text: Vec::new(),
            show_header: false,
            show_footer: false,
            show_slide_number: false,
            show_date: false,
            header_text: None,
            footer_text: None,
        }
    }

    /// Create a simple notes page with text
    #[must_use]
    pub fn simple(slide_id_ref: u32, text: &str) -> Self {
        Self::new(slide_id_ref).with_text(text)
    }

    /// Set notes text from a string
    #[must_use]
    pub fn with_text(mut self, text: &str) -> Self {
        self.text = vec![Paragraph::new(text)];
        self
    }

    /// Set notes text from paragraphs
    #[must_use]
    pub fn with_paragraphs(mut self, paragraphs: Vec<Paragraph>) -> Self {
        self.text = paragraphs;
        self
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.text.push(para);
    }

    /// Enable header with text
    #[must_use]
    pub fn with_header(mut self, text: &str) -> Self {
        self.show_header = true;
        self.header_text = Some(text.to_string());
        self
    }

    /// Enable footer with text
    #[must_use]
    pub fn with_footer(mut self, text: &str) -> Self {
        self.show_footer = true;
        self.footer_text = Some(text.to_string());
        self
    }

    /// Enable slide number
    #[must_use]
    pub fn with_slide_number(mut self) -> Self {
        self.show_slide_number = true;
        self
    }

    /// Enable date
    #[must_use]
    pub fn with_date(mut self) -> Self {
        self.show_date = true;
        self
    }

    /// Get combined text
    #[must_use]
    pub fn text_content(&self) -> String {
        self.text
            .iter()
            .map(Paragraph::text)
            .collect::<Vec<_>>()
            .join("\r")
    }

    /// Check if notes page has content
    #[must_use]
    pub fn has_content(&self) -> bool {
        !self.text.is_empty() || self.show_header || self.show_footer || self.show_slide_number
    }
}

// =============================================================================
// Notes Container Builder
// =============================================================================

/// Error type for notes operations
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesError` is the established public API name; renaming it would break downstream crates"
)]
pub type NotesError = std::io::Error;

/// Builder for Notes container record
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesContainerBuilder` is the established public API name; renaming it would break downstream crates"
)]
pub struct NotesContainerBuilder {
    notes: NotesPage,
    drawing_id: u32,
}

impl NotesContainerBuilder {
    /// Create a new builder
    #[must_use]
    pub fn new(notes: NotesPage, drawing_id: u32) -> Self {
        Self { notes, drawing_id }
    }

    /// Build the complete Notes container
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "Notes container children are small in-memory records whose total size is bounded far below u32::MAX"
    )]
    pub fn build(&self) -> Result<Vec<u8>, NotesError> {
        let mut container = Vec::new();
        let mut children = Vec::new();

        // 1) NotesAtom
        let notes_atom = self.build_notes_atom()?;
        children.push(notes_atom);

        // 2) PPDrawing with slide image and notes body placeholders
        let ppdrawing = self.build_ppdrawing()?;
        children.push(ppdrawing);

        // 3) ColorSchemeAtom
        let color_scheme = Self::build_color_scheme()?;
        children.push(color_scheme);

        // Calculate total content size
        let content_size: u32 = children.iter().map(|c| c.len() as u32).sum();

        // Write container header
        write_container_header(&mut container, record_type::NOTES, content_size)?;

        // Write children
        for child in children {
            container.extend_from_slice(&child);
        }

        Ok(container)
    }

    /// Build `NotesAtom` record
    #[allow(
        clippy::cast_possible_truncation,
        reason = "`NotesAtom::SIZE` is the constant 8, which always fits u32"
    )]
    fn build_notes_atom(&self) -> Result<Vec<u8>, NotesError> {
        let mut record = Vec::new();
        let atom = NotesAtom::new(self.notes.slide_id_ref);

        write_record_header(
            &mut record,
            0x01,
            0,
            record_type::NOTES_ATOM,
            NotesAtom::SIZE as u32,
        )?;
        record.extend_from_slice(atom.as_bytes());

        Ok(record)
    }

    /// Build `PPDrawing` with notes placeholders
    /// Uses the existing escher functions that work for slides
    fn build_ppdrawing(&self) -> Result<Vec<u8>, NotesError> {
        use super::escher::{UserShapeData, create_dg_container_with_shapes};
        use super::records::wrap_dg_into_ppdrawing;

        // Create a simple notes text shape using the same system as slides
        let mut notes_shape = UserShapeData {
            shape_type: 202, // TextBox
            ..Default::default()
        };
        // Position in EMUs (typical notes body position)
        // Convert from master units (1/576 inch) to EMUs (914400 per inch)
        notes_shape.x = ppt_master_i64_to_emu_i32(685);
        notes_shape.y = ppt_master_i64_to_emu_i32(3686);
        notes_shape.width = ppt_master_i64_to_emu_i32(4801);
        notes_shape.height = ppt_master_i64_to_emu_i32(3172);
        notes_shape.fill_color = None; // No fill for text box
        notes_shape.line_color = None; // No line

        // Set the notes text with NOTES_TYPE (2) for TextHeaderAtom
        notes_shape.text_type = 2; // NOTES_TYPE per POI TextHeaderAtom
        // Mark as notes body placeholder (type 12 per MS-PPT 2.9.39)
        notes_shape.placeholder_type = Some(12); // PT_NotesBody
        if !self.notes.text.is_empty() {
            let mut paragraphs = self.notes.text.clone();
            if paragraphs
                .iter()
                .flat_map(|paragraph| &paragraph.runs)
                .any(|run| {
                    run.international_east_asian_font_index.is_some()
                        || run.complex_script_font_index.is_some()
                })
            {
                let mut mappings = Vec::new();
                for (ordinal, run) in paragraphs
                    .iter_mut()
                    .flat_map(|paragraph| &mut paragraph.runs)
                    .enumerate()
                {
                    let run_id = u8::try_from(ordinal % 16).map_err(|_err| {
                        NotesError::new(
                            std::io::ErrorKind::InvalidInput,
                            "PP9 run identifier does not fit u8",
                        )
                    })?;
                    run.style.pp9_run_id = Some(run_id);
                    mappings.push(super::smart_tags::ShapeTextExtensionRun {
                        smart_tags: Vec::new(),
                        international_east_asian_font: run.international_east_asian_font_index,
                        complex_script_font: run.complex_script_font_index,
                    });
                }
                notes_shape.smart_tag_runs = Some(mappings);
            }
            notes_shape.paragraphs = Some(paragraphs);
        }

        let shapes = vec![notes_shape];
        let dg = create_dg_container_with_shapes(self.drawing_id, &shapes)?;
        wrap_dg_into_ppdrawing(&dg)
    }

    /// Build background `SpContainer` (per POI `PPDrawing.create()`)
    #[cfg_attr(
        not(test),
        expect(
            dead_code,
            reason = "builder mirrors POI's PPDrawing structure and is retained for pending notes-writing support"
        )
    )]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the background shape uses a fixed 8-entry property table and small in-memory records, so counts and sizes always fit u16/u32"
    )]
    fn build_background_shape(spid: u32) -> Result<Vec<u8>, NotesError> {
        let mut sp_container = Vec::new();
        let mut sp_children = Vec::new();

        // Sp record - RECT shape with BACKGROUND | HASSHAPETYPE flags
        // Per POI: sp.setOptions((short)((ShapeType.RECT.nativeId << 4) + 2))
        // RECT = 1, so options = (1 << 4) + 2 = 0x12
        let mut sp = Vec::new();
        write_escher_header(&mut sp, 0x02, 1, 0xF00A, 8)?; // version=2, instance=1 (RECT)
        sp.extend_from_slice(&spid.to_le_bytes());
        // FLAG_BACKGROUND=0x400, FLAG_HASSHAPETYPE=0x800
        sp.extend_from_slice(&0x0C00u32.to_le_bytes());
        sp_children.push(sp);

        // EscherOpt with background properties (per POI)
        let bg_props: [(u16, u32); 8] = [
            (0x0181, 0x0800_0000), // fillColor
            (0x0183, 0x0800_0005), // fillBackColor
            (0x0193, 0x0099_CCEE), // fillRectRight (approximation)
            (0x0194, 0x0076_B0DE), // fillRectBottom
            (0x01BF, 0x0012_0012), // fNoFillHitTest
            (0x01FF, 0x0008_0000), // lineBool
            (0x0304, 0x0000_0009), // blackAndWhiteMode
            (0x033F, 0x0001_0001), // fBackground
        ];
        let mut opt = Vec::new();
        write_escher_header(
            &mut opt,
            0x03,
            bg_props.len() as u16,
            0xF00B,
            (bg_props.len() * 6) as u32,
        )?;
        for (id, val) in bg_props {
            opt.extend_from_slice(&id.to_le_bytes());
            opt.extend_from_slice(&val.to_le_bytes());
        }
        sp_children.push(opt);

        let content_size: u32 = sp_children.iter().map(|c| c.len() as u32).sum();
        write_escher_header(&mut sp_container, 0x0F, 0, 0xF004, content_size)?;

        for child in sp_children {
            sp_container.extend_from_slice(&child);
        }

        Ok(sp_container)
    }

    /// Build `ColorSchemeAtom`
    fn build_color_scheme() -> Result<Vec<u8>, NotesError> {
        let mut record = Vec::new();

        // Default color scheme (same as slides)
        let colors: [u32; 8] = [
            0x00FF_FFFF, // background
            0x0000_0000, // text
            0x0080_8080, // shadow
            0x0000_0000, // title
            0x00E3_E0BB, // fill
            0x0099_3333, // accent
            0x0099_9900, // hyperlink
            0x0000_99CC, // followed hyperlink
        ];

        write_record_header(&mut record, 0x00, 1, record_type::COLOR_SCHEME_ATOM, 32)?;
        for color in colors {
            record.extend_from_slice(&color.to_le_bytes());
        }

        Ok(record)
    }
}

// =============================================================================
// Notes Collection
// =============================================================================

/// Collection of notes pages for a presentation
#[allow(
    clippy::module_name_repetitions,
    reason = "`NotesCollection` is the established public API name; renaming it would break downstream crates"
)]
#[derive(Debug, Default)]
pub struct NotesCollection {
    notes: Vec<NotesPage>,
}

impl NotesCollection {
    /// Create new empty collection
    #[must_use]
    pub fn new() -> Self {
        Self { notes: Vec::new() }
    }

    /// Add notes for a slide
    pub fn add(&mut self, notes: NotesPage) -> usize {
        let idx = self.notes.len();
        self.notes.push(notes);
        idx
    }

    /// Get notes by index
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&NotesPage> {
        self.notes.get(index)
    }

    /// Get mutable notes by index
    pub fn get_mut(&mut self, index: usize) -> Option<&mut NotesPage> {
        self.notes.get_mut(index)
    }

    /// Find notes for a slide by slide persist ID
    #[must_use]
    pub fn find_for_slide(&self, slide_id_ref: u32) -> Option<&NotesPage> {
        self.notes.iter().find(|n| n.slide_id_ref == slide_id_ref)
    }

    /// Get number of notes pages
    #[must_use]
    pub fn len(&self) -> usize {
        self.notes.len()
    }

    /// Check if empty
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.notes.is_empty()
    }

    /// Iterate over notes
    pub fn iter(&self) -> impl Iterator<Item = &NotesPage> {
        self.notes.iter()
    }

    /// Build `SlideListWithText` for notes (instance=2)
    ///
    /// # Errors
    ///
    /// Returns an error if serialization fails or the underlying writer reports an error.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "each child is a fixed-size `SlidePersistAtom`, so the total size is bounded far below u32::MAX"
    )]
    pub fn build_slide_list_with_text(
        &self,
        persist_ids: &[(u32, u32)],
    ) -> Result<Vec<u8>, NotesError> {
        if persist_ids.is_empty() {
            return Ok(Vec::new());
        }

        let mut container = Vec::new();
        let mut children = Vec::new();

        // SlidePersistAtom for each notes page
        for &(persist_id_ref, notes_identifier) in persist_ids {
            let spa = build_slide_persist_atom(persist_id_ref, notes_identifier)?;
            children.push(spa);
        }

        let content_size: u32 = children.iter().map(|c| c.len() as u32).sum();

        // SlideListWithText header with instance=2 (notes)
        write_record_header(
            &mut container,
            0x0F,
            2,
            record_type::SLIDE_LIST_WITH_TEXT,
            content_size,
        )?;

        for child in children {
            container.extend_from_slice(&child);
        }

        Ok(container)
    }
}

// =============================================================================
// Helper Functions
// =============================================================================

/// Write a PPT record header
fn write_record_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    rec_type: u16,
    length: u32,
) -> Result<(), NotesError> {
    let ver_inst = (u16::from(version) & 0x0F) | ((instance & 0x0FFF) << 4);
    writer.write_all(&ver_inst.to_le_bytes())?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Write a PPT container header
fn write_container_header<W: Write>(
    writer: &mut W,
    rec_type: u16,
    length: u32,
) -> Result<(), NotesError> {
    write_record_header(writer, 0x0F, 0, rec_type, length)
}

/// Write an Escher record header
fn write_escher_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    rec_type: u16,
    length: u32,
) -> Result<(), NotesError> {
    let ver_inst = (u16::from(version) & 0x0F) | ((instance & 0x0FFF) << 4);
    writer.write_all(&ver_inst.to_le_bytes())?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Build `SlidePersistAtom` record
fn build_slide_persist_atom(persist_id_ref: u32, identifier: u32) -> Result<Vec<u8>, NotesError> {
    let mut record = Vec::new();

    // Record header
    write_record_header(&mut record, 0x00, 0, record_type::SLIDE_PERSIST_ATOM, 20)?;

    // Atom data (20 bytes)
    record.extend_from_slice(&persist_id_ref.to_le_bytes()); // persistIdRef
    record.extend_from_slice(&0u32.to_le_bytes()); // flags
    record.extend_from_slice(&0u32.to_le_bytes()); // numberTexts
    record.extend_from_slice(&identifier.to_le_bytes()); // slideIdentifier
    record.extend_from_slice(&0u32.to_le_bytes()); // reserved

    Ok(record)
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    #[test]
    fn test_notes_atom() {
        let atom = NotesAtom::new(1);
        // Copy from packed struct to avoid unaligned access
        let slide_id_ref = { atom.slide_id_ref };
        let flags = { atom.flags };
        assert_eq!(slide_id_ref, 1);
        assert_eq!(flags, 0x0006);
    }

    #[test]
    fn test_notes_page_creation() {
        let notes = NotesPage::new(1).with_text("Speaker notes go here");

        assert_eq!(notes.slide_id_ref, 1);
        assert_eq!(notes.text.len(), 1);
        assert!(notes.has_content());
    }

    #[test]
    fn test_notes_collection() {
        let mut collection = NotesCollection::new();
        collection.add(NotesPage::new(1).with_text("Notes 1"));
        collection.add(NotesPage::new(2).with_text("Notes 2"));

        assert_eq!(collection.len(), 2);
        assert!(collection.find_for_slide(1).is_some());
    }

    #[test]
    fn test_notes_container_builder() {
        let notes = NotesPage::new(1).with_text("Test notes");
        let builder = NotesContainerBuilder::new(notes, 3);
        let container = builder.build().unwrap();

        // Should produce non-empty container
        assert!(!container.is_empty());

        // Check container starts with Notes record type
        let rec_type = u16::from_le_bytes([container[2], container[3]]);
        assert_eq!(rec_type, record_type::NOTES);
    }

    #[test]
    fn test_notes_background_uses_shape_boolean_group() {
        let background = NotesContainerBuilder::build_background_shape(0x0C04).unwrap();
        let expected_properties = [
            (0x0181u16, 0x0800_0000u32),
            (0x0183, 0x0800_0005),
            (0x0193, 0x0099_CCEE),
            (0x0194, 0x0076_B0DE),
            (0x01BF, 0x0012_0012),
            (0x01FF, 0x0008_0000),
            (0x0304, 0x0000_0009),
            (0x033F, 0x0001_0001),
        ];
        let mut expected_bytes = Vec::new();
        for (id, value) in expected_properties {
            expected_bytes.extend_from_slice(&id.to_le_bytes());
            expected_bytes.extend_from_slice(&value.to_le_bytes());
        }
        assert!(
            background
                .windows(expected_bytes.len())
                .any(|window| window == expected_bytes),
            "notes background must contain the exact OfficeArt property table"
        );
    }
}
