//! Notes slide support for PPT files
//!
//! This module handles speaker notes pages in PowerPoint presentations.
//! Notes pages contain the slide thumbnail and notes text area.
//!
//! Reference: [MS-PPT] Section 2.5.5 - NotesContainer

use std::io::Write;
use zerocopy::IntoBytes;
use zerocopy_derive::*;

use super::text_format::Paragraph;

// =============================================================================
// PPT Record Types for Notes
// =============================================================================

/// Record types for notes
pub mod record_type {
    /// Notes container (RT_Notes)
    pub const NOTES: u16 = 0x03F0;
    /// NotesAtom
    pub const NOTES_ATOM: u16 = 0x03F1;
    /// PPDrawing (contains Escher records)
    pub const PP_DRAWING: u16 = 0x040C;
    /// TextHeader (type of text)
    pub const TEXT_HEADER_ATOM: u16 = 0x0F9F;
    /// TextCharsAtom (UTF-16 text)
    pub const TEXT_CHARS_ATOM: u16 = 0x0FA0;
    /// TextBytesAtom (ASCII text)
    pub const TEXT_BYTES_ATOM: u16 = 0x0FA8;
    /// StyleTextPropAtom
    pub const STYLE_TEXT_PROP_ATOM: u16 = 0x0FA1;
    /// ColorSchemeAtom
    pub const COLOR_SCHEME_ATOM: u16 = 0x07F0;
    /// SlideListWithText for notes (instance=2)
    pub const SLIDE_LIST_WITH_TEXT: u16 = 0x0FF0;
    /// SlidePersistAtom
    pub const SLIDE_PERSIST_ATOM: u16 = 0x03F3;
    /// TextRuler
    pub const TEXT_RULER_ATOM: u16 = 0x0FA2;
    /// NotesTextViewInfo
    pub const NOTES_TEXT_VIEW_INFO: u16 = 0x03F5;
}

// =============================================================================
// NotesAtom (MS-PPT 2.5.5.1)
// =============================================================================

/// NotesAtom structure (8 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct NotesAtom {
    /// Slide persist ID reference
    pub slide_id_ref: u32,
    /// Notes flags
    pub flags: u16,
    /// Reserved
    pub reserved: u16,
}

impl NotesAtom {
    /// Size of the atom data
    pub const SIZE: usize = 8;

    /// Create a new NotesAtom
    pub fn new(slide_id_ref: u32) -> Self {
        Self {
            slide_id_ref,
            flags: 0x0006, // fFollowMasterObjects | fFollowMasterScheme
            reserved: 0,
        }
    }
}

/// Notes flags
pub mod notes_flags {
    /// Follow master objects
    pub const FOLLOW_MASTER_OBJECTS: u16 = 0x0001;
    /// Follow master color scheme
    pub const FOLLOW_MASTER_SCHEME: u16 = 0x0002;
    /// Follow master background
    pub const FOLLOW_MASTER_BACKGROUND: u16 = 0x0004;
}

// =============================================================================
// Placeholder Types for Notes
// =============================================================================

/// Placeholder types used in notes pages
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum NotesPlaceholderType {
    /// Slide image placeholder
    SlideImage = 0x05,
    /// Notes body placeholder
    NotesBody = 0x06,
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
    pub fn simple(slide_id_ref: u32, text: &str) -> Self {
        Self::new(slide_id_ref).with_text(text)
    }

    /// Set notes text from a string
    pub fn with_text(mut self, text: &str) -> Self {
        self.text = vec![Paragraph::new(text)];
        self
    }

    /// Set notes text from paragraphs
    pub fn with_paragraphs(mut self, paragraphs: Vec<Paragraph>) -> Self {
        self.text = paragraphs;
        self
    }

    /// Add a paragraph
    pub fn add_paragraph(&mut self, para: Paragraph) {
        self.text.push(para);
    }

    /// Enable header with text
    pub fn with_header(mut self, text: &str) -> Self {
        self.show_header = true;
        self.header_text = Some(text.to_string());
        self
    }

    /// Enable footer with text
    pub fn with_footer(mut self, text: &str) -> Self {
        self.show_footer = true;
        self.footer_text = Some(text.to_string());
        self
    }

    /// Enable slide number
    pub fn with_slide_number(mut self) -> Self {
        self.show_slide_number = true;
        self
    }

    /// Enable date
    pub fn with_date(mut self) -> Self {
        self.show_date = true;
        self
    }

    /// Get combined text
    pub fn text_content(&self) -> String {
        self.text
            .iter()
            .map(|p| p.text())
            .collect::<Vec<_>>()
            .join("\r")
    }

    /// Check if notes page has content
    pub fn has_content(&self) -> bool {
        !self.text.is_empty() || self.show_header || self.show_footer || self.show_slide_number
    }
}

// =============================================================================
// Notes Container Builder
// =============================================================================

/// Error type for notes operations
pub type NotesError = std::io::Error;

/// Builder for Notes container record
pub struct NotesContainerBuilder {
    notes: NotesPage,
    drawing_id: u32,
}

impl NotesContainerBuilder {
    /// Create a new builder
    pub fn new(notes: NotesPage, drawing_id: u32) -> Self {
        Self { notes, drawing_id }
    }

    /// Build the complete Notes container
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
        let color_scheme = self.build_color_scheme()?;
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

    /// Build NotesAtom record
    fn build_notes_atom(&self) -> Result<Vec<u8>, NotesError> {
        let mut record = Vec::new();
        let atom = NotesAtom::new(self.notes.slide_id_ref);

        write_record_header(
            &mut record,
            0x02,
            0,
            record_type::NOTES_ATOM,
            NotesAtom::SIZE as u32,
        )?;
        record.extend_from_slice(atom.as_bytes());

        Ok(record)
    }

    /// Build PPDrawing with notes placeholders
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
        let master_to_emu = |m: i64| -> i32 { (m * 914400 / 576) as i32 };
        notes_shape.x = master_to_emu(685);
        notes_shape.y = master_to_emu(3686);
        notes_shape.width = master_to_emu(4801);
        notes_shape.height = master_to_emu(3172);
        notes_shape.fill_color = None; // No fill for text box
        notes_shape.line_color = None; // No line

        // Set the notes text with NOTES_TYPE (2) for TextHeaderAtom
        notes_shape.text_type = 2; // NOTES_TYPE per POI TextHeaderAtom
        // Mark as notes body placeholder (type 12 per MS-PPT 2.9.39)
        notes_shape.placeholder_type = Some(12); // PT_NotesBody
        if !self.notes.text.is_empty() {
            notes_shape.text = Some(self.notes.text_content());
        }

        let shapes = vec![notes_shape];
        let dg = create_dg_container_with_shapes(self.drawing_id, &shapes)?;
        wrap_dg_into_ppdrawing(&dg)
    }

    /// Build DgContainer with notes shapes
    /// Per POI PPDrawing.create(): DgContainer contains DgRecord, SpgrContainer, and background SpContainer
    #[allow(dead_code)]
    fn build_dg_container(&self) -> Result<Vec<u8>, NotesError> {
        let mut dg_container = Vec::new();
        let mut dg_children = Vec::new();

        // User shapes count (not including patriarch): slide image + notes body = 2
        let user_shape_count = 2u32;
        // POI: numShapes = user shapes + 1 (for background) but background doesn't count in PPDrawing.create()
        // Actually POI sets numShapes=1 initially, then increments when shapes are added
        // For notes, let's count: patriarch(1) + slide_image + notes_body = 3 total
        let shape_count = 1 + user_shape_count;
        let base_spid = self.drawing_id << 10;
        let last_spid = base_spid + shape_count;

        // EscherDg (drawing info) - version=0, instance=drawing_id
        let mut dg = Vec::new();
        write_escher_header(&mut dg, 0x00, self.drawing_id as u16, 0xF008, 8)?;
        dg.extend_from_slice(&shape_count.to_le_bytes());
        dg.extend_from_slice(&last_spid.to_le_bytes());
        dg_children.push(dg);

        // SpgrContainer with patriarch and user shapes
        let spgr_content = self.build_spgr_container()?;
        let mut spgr = Vec::new();
        write_escher_header(&mut spgr, 0x0F, 0, 0xF003, spgr_content.len() as u32)?;
        spgr.extend_from_slice(&spgr_content);
        dg_children.push(spgr);

        // Background SpContainer (per POI PPDrawing.create() - directly in DgContainer)
        let bg_sp = self.build_background_shape(base_spid + shape_count + 1)?;
        dg_children.push(bg_sp);

        // Calculate size
        let content_size: u32 = dg_children.iter().map(|c| c.len() as u32).sum();

        // DgContainer header
        write_escher_header(&mut dg_container, 0x0F, 0, 0xF002, content_size)?;

        for child in dg_children {
            dg_container.extend_from_slice(&child);
        }

        Ok(dg_container)
    }

    /// Build background SpContainer (per POI PPDrawing.create())
    #[allow(dead_code)]
    fn build_background_shape(&self, spid: u32) -> Result<Vec<u8>, NotesError> {
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
            (0x0181, 0x08000000), // fillColor
            (0x0183, 0x08000005), // fillBackColor
            (0x0185, 0x0099CCEE), // fillRectRight (approximation)
            (0x0186, 0x0076B0DE), // fillRectBottom
            (0x01BF, 0x00120012), // fNoFillHitTest
            (0x01FF, 0x00080000), // lineBool
            (0x03BF, 0x00000009), // shadowBool
            (0x03FF, 0x00010001), // shapeBool - BACKGROUND_SHAPE
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

    /// Build SpgrContainer with group and placeholder shapes
    #[allow(dead_code)]
    fn build_spgr_container(&self) -> Result<Vec<u8>, NotesError> {
        let mut content = Vec::new();
        let base_spid = self.drawing_id << 10;

        // Group patriarch
        let group = self.build_group_patriarch(base_spid)?;
        content.extend_from_slice(&group);

        // Slide image placeholder
        let slide_image = self.build_slide_image_placeholder(base_spid + 1)?;
        content.extend_from_slice(&slide_image);

        // Notes body placeholder
        let notes_body = self.build_notes_body_placeholder(base_spid + 2)?;
        content.extend_from_slice(&notes_body);

        Ok(content)
    }

    /// Build group patriarch SpContainer
    #[allow(dead_code)]
    fn build_group_patriarch(&self, spid: u32) -> Result<Vec<u8>, NotesError> {
        let mut sp_container = Vec::new();
        let mut sp_children = Vec::new();

        // Spgr record (group coords - all zeros)
        let mut spgr = Vec::new();
        write_escher_header(&mut spgr, 0x01, 0, 0xF009, 16)?;
        spgr.extend_from_slice(&[0u8; 16]);
        sp_children.push(spgr);

        // Sp record (group shape)
        let mut sp = Vec::new();
        write_escher_header(&mut sp, 0x02, 0, 0xF00A, 8)?;
        sp.extend_from_slice(&spid.to_le_bytes());
        sp.extend_from_slice(&0x0005u32.to_le_bytes()); // fGroup | fPatriarch
        sp_children.push(sp);

        let content_size: u32 = sp_children.iter().map(|c| c.len() as u32).sum();
        write_escher_header(&mut sp_container, 0x0F, 0, 0xF004, content_size)?;

        for child in sp_children {
            sp_container.extend_from_slice(&child);
        }

        Ok(sp_container)
    }

    /// Build slide image placeholder SpContainer
    #[allow(dead_code)]
    fn build_slide_image_placeholder(&self, spid: u32) -> Result<Vec<u8>, NotesError> {
        self.build_placeholder_shape(
            spid,
            NotesPlaceholderType::SlideImage,
            (685, 576, 5486, 3514), // Typical notes slide image position
        )
    }

    /// Build notes body placeholder SpContainer with text
    #[allow(dead_code)]
    fn build_notes_body_placeholder(&self, spid: u32) -> Result<Vec<u8>, NotesError> {
        let mut sp_container = Vec::new();
        let mut sp_children = Vec::new();
        let anchor = (685u16, 3686u16, 5486u16, 6858u16); // Below slide image

        // Sp record (TextBox shape = 202)
        let mut sp = Vec::new();
        write_escher_header(&mut sp, 0x02, 202, 0xF00A, 8)?;
        sp.extend_from_slice(&spid.to_le_bytes());
        sp.extend_from_slice(&0x0A00u32.to_le_bytes()); // fHaveAnchor | fHaveSpt
        sp_children.push(sp);

        // Opt record (properties)
        let props: [(u16, u32); 4] = [
            (0x007F, 0x00010005), // lockAggr
            (0x0181, 0x08000004), // fillColor (scheme fill)
            (0x01BF, 0x00010001), // fNoFillHitTest (filled)
            (0x01FF, 0x00090001), // shapeBool
        ];
        let mut opt = Vec::new();
        write_escher_header(
            &mut opt,
            0x03,
            props.len() as u16,
            0xF00B,
            (props.len() * 6) as u32,
        )?;
        for (id, val) in props {
            opt.extend_from_slice(&id.to_le_bytes());
            opt.extend_from_slice(&val.to_le_bytes());
        }
        sp_children.push(opt);

        // ClientAnchor (position)
        let mut client_anchor = Vec::new();
        write_escher_header(&mut client_anchor, 0x00, 0, 0xF010, 8)?;
        client_anchor.extend_from_slice(&anchor.0.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.1.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.2.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.3.to_le_bytes());
        sp_children.push(client_anchor);

        // ClientTextbox with notes text (if any)
        if !self.notes.text.is_empty() {
            let textbox = self.build_client_textbox()?;
            sp_children.push(textbox);
        }

        // ClientData with OEPlaceholderAtom
        let mut client_data = Vec::new();
        let oe_placeholder = build_oe_placeholder_atom(0, NotesPlaceholderType::NotesBody as u8)?;
        write_escher_header(
            &mut client_data,
            0x0F,
            0,
            0xF011,
            oe_placeholder.len() as u32,
        )?;
        client_data.extend_from_slice(&oe_placeholder);
        sp_children.push(client_data);

        let content_size: u32 = sp_children.iter().map(|c| c.len() as u32).sum();
        write_escher_header(&mut sp_container, 0x0F, 0, 0xF004, content_size)?;

        for child in sp_children {
            sp_container.extend_from_slice(&child);
        }

        Ok(sp_container)
    }

    /// Build ClientTextbox record with notes text
    #[allow(dead_code)]
    fn build_client_textbox(&self) -> Result<Vec<u8>, NotesError> {
        let mut textbox = Vec::new();
        let mut children = Vec::new();

        // Get combined text content
        let text_content = self.notes.text_content();

        // TextHeaderAtom (type=3999) - Per POI TextHeaderAtom.NOTES_TYPE = 2
        let mut text_header = Vec::new();
        write_record_header(&mut text_header, 0x00, 0, record_type::TEXT_HEADER_ATOM, 4)?;
        text_header.extend_from_slice(&2u32.to_le_bytes()); // txType = NOTES_TYPE (2)
        children.push(text_header);

        // TextCharsAtom (type=4000) - UTF-16LE text
        let utf16: Vec<u16> = text_content.encode_utf16().collect();
        let text_len = (utf16.len() * 2) as u32;
        let mut text_chars = Vec::new();
        write_record_header(
            &mut text_chars,
            0x00,
            0,
            record_type::TEXT_CHARS_ATOM,
            text_len,
        )?;
        for ch in utf16 {
            text_chars.extend_from_slice(&ch.to_le_bytes());
        }
        children.push(text_chars);

        // Calculate total size
        let content_size: u32 = children.iter().map(|c| c.len() as u32).sum();

        // ClientTextbox Escher header (0xF00D)
        write_escher_header(&mut textbox, 0x00, 0, 0xF00D, content_size)?;

        for child in children {
            textbox.extend_from_slice(&child);
        }

        Ok(textbox)
    }

    /// Build a placeholder shape SpContainer (without text)
    #[allow(dead_code)]
    fn build_placeholder_shape(
        &self,
        spid: u32,
        placeholder_type: NotesPlaceholderType,
        anchor: (u16, u16, u16, u16),
    ) -> Result<Vec<u8>, NotesError> {
        let mut sp_container = Vec::new();
        let mut sp_children = Vec::new();

        // Sp record (TextBox shape = 202)
        let mut sp = Vec::new();
        write_escher_header(&mut sp, 0x02, 202, 0xF00A, 8)?;
        sp.extend_from_slice(&spid.to_le_bytes());
        sp.extend_from_slice(&0x0A00u32.to_le_bytes()); // fHaveAnchor | fHaveSpt
        sp_children.push(sp);

        // Opt record (properties)
        let props: [(u16, u32); 4] = [
            (0x007F, 0x00010005), // lockAggr
            (0x0181, 0x08000004), // fillColor (scheme fill)
            (0x01BF, 0x00010001), // fNoFillHitTest (filled)
            (0x01FF, 0x00090001), // shapeBool
        ];
        let mut opt = Vec::new();
        write_escher_header(
            &mut opt,
            0x03,
            props.len() as u16,
            0xF00B,
            (props.len() * 6) as u32,
        )?;
        for (id, val) in props {
            opt.extend_from_slice(&id.to_le_bytes());
            opt.extend_from_slice(&val.to_le_bytes());
        }
        sp_children.push(opt);

        // ClientAnchor (position)
        let mut client_anchor = Vec::new();
        write_escher_header(&mut client_anchor, 0x00, 0, 0xF010, 8)?;
        client_anchor.extend_from_slice(&anchor.0.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.1.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.2.to_le_bytes());
        client_anchor.extend_from_slice(&anchor.3.to_le_bytes());
        sp_children.push(client_anchor);

        // ClientData with OEPlaceholderAtom
        let mut client_data = Vec::new();
        let oe_placeholder = build_oe_placeholder_atom(0, placeholder_type as u8)?;
        write_escher_header(
            &mut client_data,
            0x0F,
            0,
            0xF011,
            oe_placeholder.len() as u32,
        )?;
        client_data.extend_from_slice(&oe_placeholder);
        sp_children.push(client_data);

        let content_size: u32 = sp_children.iter().map(|c| c.len() as u32).sum();
        write_escher_header(&mut sp_container, 0x0F, 0, 0xF004, content_size)?;

        for child in sp_children {
            sp_container.extend_from_slice(&child);
        }

        Ok(sp_container)
    }

    /// Build ColorSchemeAtom
    fn build_color_scheme(&self) -> Result<Vec<u8>, NotesError> {
        let mut record = Vec::new();

        // Default color scheme (same as slides)
        let colors: [u32; 8] = [
            0x00FFFFFF, // background
            0x00000000, // text
            0x00808080, // shadow
            0x00000000, // title
            0x00E3E0BB, // fill
            0x00993333, // accent
            0x00999900, // hyperlink
            0x000099CC, // followed hyperlink
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
#[derive(Debug, Default)]
pub struct NotesCollection {
    notes: Vec<NotesPage>,
}

impl NotesCollection {
    /// Create new empty collection
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
    pub fn get(&self, index: usize) -> Option<&NotesPage> {
        self.notes.get(index)
    }

    /// Get mutable notes by index
    pub fn get_mut(&mut self, index: usize) -> Option<&mut NotesPage> {
        self.notes.get_mut(index)
    }

    /// Find notes for a slide by slide persist ID
    pub fn find_for_slide(&self, slide_id_ref: u32) -> Option<&NotesPage> {
        self.notes.iter().find(|n| n.slide_id_ref == slide_id_ref)
    }

    /// Get number of notes pages
    pub fn len(&self) -> usize {
        self.notes.len()
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.notes.is_empty()
    }

    /// Iterate over notes
    pub fn iter(&self) -> impl Iterator<Item = &NotesPage> {
        self.notes.iter()
    }

    /// Build SlideListWithText for notes (instance=2)
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
    let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
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
#[allow(dead_code)]
fn write_escher_header<W: Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    rec_type: u16,
    length: u32,
) -> Result<(), NotesError> {
    let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
    writer.write_all(&ver_inst.to_le_bytes())?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Build OEPlaceholderAtom PPT record
#[allow(dead_code)]
fn build_oe_placeholder_atom(position: u32, placeholder_type: u8) -> Result<Vec<u8>, NotesError> {
    let mut record = Vec::new();

    // Record header (8 bytes)
    write_record_header(&mut record, 0x00, 0, 0x0BC3, 8)?;

    // Atom data (8 bytes)
    record.extend_from_slice(&position.to_le_bytes()); // placementId
    record.push(placeholder_type); // placeholderType
    record.push(0x01); // placeholderSize = quarter
    record.extend_from_slice(&[0x00, 0x00]); // unused

    Ok(record)
}

/// Build SlidePersistAtom record
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
}
