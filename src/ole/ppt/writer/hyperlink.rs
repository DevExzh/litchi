//! Hyperlink support for PPT files
//!
//! This module handles hyperlinks in PowerPoint presentations including
//! URL links, internal slide links, and action buttons.
//!
//! Reference: [MS-PPT] Section 2.8 - Interactive Information

use zerocopy::IntoBytes;
use zerocopy_derive::*;

// =============================================================================
// PPT Record Types for Hyperlinks
// =============================================================================

/// Record types for interactive/hyperlink records
pub mod record_type {
    /// InteractiveInfo container (RT_InteractiveInfo)
    pub const INTERACTIVE_INFO: u16 = 0x0FF2;
    /// InteractiveInfoAtom
    pub const INTERACTIVE_INFO_ATOM: u16 = 0x0FF3;
    /// ExHyperlink container
    pub const EX_HYPERLINK: u16 = 0x0FD7;
    /// ExHyperlinkAtom
    pub const EX_HYPERLINK_ATOM: u16 = 0x0FD3;
    /// CString (unicode string)
    pub const CSTRING: u16 = 0x0FBA;
    /// ExObjList container
    pub const EX_OBJ_LIST: u16 = 0x0409;
    /// ExObjListAtom
    pub const EX_OBJ_LIST_ATOM: u16 = 0x040A;
    /// MouseClick in shape
    pub const MOUSE_CLICK: u16 = 0x0001; // instance value
    /// MouseOver in shape
    pub const MOUSE_OVER: u16 = 0x0000; // instance value
}

// =============================================================================
// Hyperlink Action Types
// =============================================================================

/// Action type for hyperlinks (per POI InteractiveInfoAtom)
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum HyperlinkAction {
    /// No action (ACTION_NONE)
    #[default]
    None = 0x00,
    /// Macro (ACTION_MACRO)
    Macro = 0x01,
    /// Run program (ACTION_RUNPROGRAM)
    RunProgram = 0x02,
    /// Jump to slide (ACTION_JUMP) - use with JumpAction
    Jump = 0x03,
    /// Go to hyperlink URL (ACTION_HYPERLINK)
    Hyperlink = 0x04,
    /// OLE action (ACTION_OLE)
    OleAction = 0x05,
    /// Media (ACTION_MEDIA)
    Media = 0x06,
    /// Custom show (ACTION_CUSTOMSHOW)
    CustomShow = 0x07,
}

/// Jump action flags
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum JumpAction {
    /// No jump
    #[default]
    None = 0x00,
    /// Jump to next slide
    NextSlide = 0x01,
    /// Jump to previous slide
    PreviousSlide = 0x02,
    /// Jump to first slide
    FirstSlide = 0x03,
    /// Jump to last slide
    LastSlide = 0x04,
    /// Jump to last viewed slide
    LastViewed = 0x05,
    /// End show
    EndShow = 0x06,
}

// =============================================================================
// InteractiveInfoAtom (MS-PPT 2.8.1)
// =============================================================================

/// InteractiveInfoAtom structure (16 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct InteractiveInfoAtom {
    /// Sound reference (0 if none)
    pub sound_ref: u32,
    /// ExHyperlink reference (0 if none)  
    pub hyperlink_ref: u32,
    /// Action type
    pub action: u8,
    /// OLE verb (0 if not OLE)
    pub ole_verb: u8,
    /// Jump action
    pub jump: u8,
    /// Flags
    pub flags: u8,
    /// Hyperlink type (0=default, 1=external)
    pub hyperlink_type: u8,
    /// Reserved bytes
    pub reserved: [u8; 3],
}

impl InteractiveInfoAtom {
    /// Size of the atom data
    pub const SIZE: usize = 16;

    /// Create a new InteractiveInfoAtom for URL hyperlink (per POI linkToUrl)
    pub fn url_link(hyperlink_id: u32) -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: hyperlink_id,
            action: HyperlinkAction::Hyperlink as u8, // ACTION_HYPERLINK = 4
            ole_verb: 0,
            jump: JumpAction::None as u8, // JUMP_NONE = 0
            flags: 0x04,                  // fAnimated
            hyperlink_type: 0x08,         // LINK_Url
            reserved: [0; 3],
        }
    }

    /// Create a new InteractiveInfoAtom for slide number link
    pub fn slide_link(hyperlink_id: u32) -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: hyperlink_id,
            action: HyperlinkAction::Hyperlink as u8, // ACTION_HYPERLINK for specific slide
            ole_verb: 0,
            jump: JumpAction::None as u8,
            flags: 0x04,
            hyperlink_type: 0x07, // LINK_SlideNumber
            reserved: [0; 3],
        }
    }

    /// Create atom for next slide action (per POI linkToNextSlide)
    pub fn next_slide() -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: 0,
            action: HyperlinkAction::Jump as u8, // ACTION_JUMP
            ole_verb: 0,
            jump: JumpAction::NextSlide as u8, // JUMP_NEXTSLIDE
            flags: 0x04,
            hyperlink_type: 0x00, // LINK_NextSlide
            reserved: [0; 3],
        }
    }

    /// Create atom for previous slide action
    pub fn prev_slide() -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: 0,
            action: HyperlinkAction::Jump as u8, // ACTION_JUMP
            ole_verb: 0,
            jump: JumpAction::PreviousSlide as u8, // JUMP_PREVIOUSSLIDE
            flags: 0x04,
            hyperlink_type: 0x01, // LINK_PreviousSlide
            reserved: [0; 3],
        }
    }

    /// Create atom for first slide action
    pub fn first_slide() -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: 0,
            action: HyperlinkAction::Jump as u8,
            ole_verb: 0,
            jump: JumpAction::FirstSlide as u8,
            flags: 0x04,
            hyperlink_type: 0x02, // LINK_FirstSlide
            reserved: [0; 3],
        }
    }

    /// Create atom for last slide action
    pub fn last_slide() -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: 0,
            action: HyperlinkAction::Jump as u8,
            ole_verb: 0,
            jump: JumpAction::LastSlide as u8,
            flags: 0x04,
            hyperlink_type: 0x03, // LINK_LastSlide
            reserved: [0; 3],
        }
    }

    /// Create atom for end show action
    pub fn end_show() -> Self {
        Self {
            sound_ref: 0,
            hyperlink_ref: 0,
            action: HyperlinkAction::Jump as u8,
            ole_verb: 0,
            jump: JumpAction::EndShow as u8,
            flags: 0x04,
            hyperlink_type: 0xFF, // LINK_NULL for end show
            reserved: [0; 3],
        }
    }
}

// =============================================================================
// ExHyperlinkAtom (MS-PPT 2.10.18)
// =============================================================================

/// ExHyperlinkAtom structure (4 bytes)
#[derive(Debug, Clone, Copy, FromBytes, IntoBytes, Immutable, KnownLayout)]
#[repr(C, packed)]
pub struct ExHyperlinkAtom {
    /// Hyperlink ID (1-based)
    pub hyperlink_id: u32,
}

// =============================================================================
// Hyperlink Definition
// =============================================================================

/// Hyperlink target types
#[derive(Debug, Clone)]
pub enum HyperlinkTarget {
    /// URL (external web link)
    Url(String),
    /// File path
    File(String),
    /// Slide number (1-based)
    Slide(u32),
    /// Next slide
    NextSlide,
    /// Previous slide
    PrevSlide,
    /// First slide
    FirstSlide,
    /// Last slide
    LastSlide,
    /// End show
    EndShow,
    /// Custom show by name
    CustomShow(String),
}

/// Complete hyperlink definition
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Unique ID (assigned by HyperlinkCollection)
    pub id: u32,
    /// Display text (tooltip)
    pub display_text: Option<String>,
    /// Target
    pub target: HyperlinkTarget,
    /// Target frame (for URLs)
    pub target_frame: Option<String>,
}

impl Hyperlink {
    /// Create URL hyperlink
    pub fn url(url: impl Into<String>) -> Self {
        Self {
            id: 0,
            display_text: None,
            target: HyperlinkTarget::Url(url.into()),
            target_frame: None,
        }
    }

    /// Create file hyperlink
    pub fn file(path: impl Into<String>) -> Self {
        Self {
            id: 0,
            display_text: None,
            target: HyperlinkTarget::File(path.into()),
            target_frame: None,
        }
    }

    /// Create slide hyperlink
    pub fn slide(slide_num: u32) -> Self {
        Self {
            id: 0,
            display_text: None,
            target: HyperlinkTarget::Slide(slide_num),
            target_frame: None,
        }
    }

    /// Create next slide hyperlink
    pub fn next_slide() -> Self {
        Self {
            id: 0,
            display_text: None,
            target: HyperlinkTarget::NextSlide,
            target_frame: None,
        }
    }

    /// Create previous slide hyperlink
    pub fn prev_slide() -> Self {
        Self {
            id: 0,
            display_text: None,
            target: HyperlinkTarget::PrevSlide,
            target_frame: None,
        }
    }

    /// Set display text
    pub fn with_display_text(mut self, text: impl Into<String>) -> Self {
        self.display_text = Some(text.into());
        self
    }

    /// Set target frame
    pub fn with_target_frame(mut self, frame: impl Into<String>) -> Self {
        self.target_frame = Some(frame.into());
        self
    }

    /// Check if this is an external link (URL or file)
    pub fn is_external(&self) -> bool {
        matches!(
            self.target,
            HyperlinkTarget::Url(_) | HyperlinkTarget::File(_)
        )
    }

    /// Get the URL/path string if applicable
    pub fn target_string(&self) -> Option<&str> {
        match &self.target {
            HyperlinkTarget::Url(s) | HyperlinkTarget::File(s) => Some(s.as_str()),
            HyperlinkTarget::CustomShow(s) => Some(s.as_str()),
            _ => None,
        }
    }

    /// Build InteractiveInfoAtom for this hyperlink
    pub fn build_interactive_info_atom(&self) -> InteractiveInfoAtom {
        match &self.target {
            HyperlinkTarget::Url(_) | HyperlinkTarget::File(_) => {
                InteractiveInfoAtom::url_link(self.id)
            },
            HyperlinkTarget::Slide(_) => InteractiveInfoAtom::slide_link(self.id),
            HyperlinkTarget::NextSlide => InteractiveInfoAtom::next_slide(),
            HyperlinkTarget::PrevSlide => InteractiveInfoAtom::prev_slide(),
            HyperlinkTarget::FirstSlide => InteractiveInfoAtom::first_slide(),
            HyperlinkTarget::LastSlide => InteractiveInfoAtom::last_slide(),
            HyperlinkTarget::EndShow => InteractiveInfoAtom::end_show(),
            HyperlinkTarget::CustomShow(_) => InteractiveInfoAtom::slide_link(self.id),
        }
    }
}

// =============================================================================
// Hyperlink Collection
// =============================================================================

/// Collection of hyperlinks for a presentation
#[derive(Debug, Default)]
pub struct HyperlinkCollection {
    hyperlinks: Vec<Hyperlink>,
    next_id: u32,
}

impl HyperlinkCollection {
    /// Create new empty collection
    pub fn new() -> Self {
        Self {
            hyperlinks: Vec::new(),
            next_id: 1,
        }
    }

    /// Add a hyperlink and return its ID
    pub fn add(&mut self, mut hyperlink: Hyperlink) -> u32 {
        let id = self.next_id;
        self.next_id += 1;
        hyperlink.id = id;
        self.hyperlinks.push(hyperlink);
        id
    }

    /// Get hyperlink by ID
    pub fn get(&self, id: u32) -> Option<&Hyperlink> {
        self.hyperlinks.iter().find(|h| h.id == id)
    }

    /// Get number of hyperlinks
    pub fn len(&self) -> usize {
        self.hyperlinks.len()
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.hyperlinks.is_empty()
    }

    /// Iterate over hyperlinks
    pub fn iter(&self) -> impl Iterator<Item = &Hyperlink> {
        self.hyperlinks.iter()
    }

    /// Build ExObjList container for document
    pub fn build_ex_obj_list(&self) -> Result<Vec<u8>, std::io::Error> {
        if self.hyperlinks.is_empty() {
            return Ok(Vec::new());
        }

        let mut container = Vec::new();

        // Build ExHyperlink records for external links
        let mut children = Vec::new();

        // ExObjListAtom
        let mut list_atom = Vec::new();
        write_ppt_header(&mut list_atom, record_type::EX_OBJ_LIST_ATOM, 4)?;
        list_atom.extend_from_slice(&(self.next_id - 1).to_le_bytes()); // Object seed
        children.push(list_atom);

        // ExHyperlink for ALL hyperlinks (per POI, all hyperlinks need ExHyperlink)
        for hyperlink in &self.hyperlinks {
            let ex_hyperlink = self.build_ex_hyperlink(hyperlink)?;
            children.push(ex_hyperlink);
        }

        // Calculate total size
        let content_size: u32 = children.iter().map(|c| c.len() as u32).sum();

        // Write container header
        write_ppt_container_header(&mut container, record_type::EX_OBJ_LIST, content_size)?;

        // Write children
        for child in children {
            container.extend_from_slice(&child);
        }

        Ok(container)
    }

    /// Build ExHyperlink container for a single hyperlink
    /// Per POI ExHyperlink: contains ExHyperlinkAtom + 2 CStrings (title, URL)
    fn build_ex_hyperlink(&self, hyperlink: &Hyperlink) -> Result<Vec<u8>, std::io::Error> {
        let mut container = Vec::new();
        let mut children = Vec::new();

        // ExHyperlinkAtom
        let mut atom = Vec::new();
        write_ppt_header(&mut atom, record_type::EX_HYPERLINK_ATOM, 4)?;
        atom.extend_from_slice(&hyperlink.id.to_le_bytes());
        children.push(atom);

        // Get title and URL based on hyperlink type (per POI HSLFHyperlink)
        let (title, url, link_options) = match &hyperlink.target {
            HyperlinkTarget::Url(u) => {
                let t = hyperlink.display_text.as_deref().unwrap_or(u);
                (t.to_string(), u.clone(), 0x10u16) // URL links: options=0x10
            },
            HyperlinkTarget::File(f) => {
                let t = hyperlink.display_text.as_deref().unwrap_or(f);
                (t.to_string(), f.clone(), 0x10u16)
            },
            HyperlinkTarget::Slide(num) => {
                // Per POI: linkToDocument(sheetNumber, slideNumber, alias, 0x30)
                // URL format: "sheetNumber,slideNumber,alias"
                let alias = format!("Slide {}", num);
                let url = format!("1,{},{}", num, alias); // sheetNumber=1 for main presentation
                (alias.clone(), url, 0x30u16) // Slide links: options=0x30
            },
            HyperlinkTarget::NextSlide => {
                // Per POI: linkToDocument(1, -1, "NEXT", 0x10)
                ("NEXT".to_string(), "1,-1,NEXT".to_string(), 0x10u16)
            },
            HyperlinkTarget::PrevSlide => ("PREV".to_string(), "1,-1,PREV".to_string(), 0x10u16),
            HyperlinkTarget::FirstSlide => ("FIRST".to_string(), "1,-1,FIRST".to_string(), 0x10u16),
            HyperlinkTarget::LastSlide => ("LAST".to_string(), "1,-1,LAST".to_string(), 0x10u16),
            HyperlinkTarget::EndShow => {
                ("End Show".to_string(), "1,-1,End Show".to_string(), 0x10u16)
            },
            HyperlinkTarget::CustomShow(name) => (name.clone(), name.clone(), 0x10u16),
        };

        // CString records per POI ExHyperlink structure:
        // 1. linkDetailsA (title) with options=0x00 (instance=0)
        // 2. linkDetailsB (URL) with options=link_options
        let title_cstring = build_cstring_with_options(0x00, &title)?;
        children.push(title_cstring);

        let url_cstring = build_cstring_with_options(link_options, &url)?;
        children.push(url_cstring);

        // Calculate total size
        let content_size: u32 = children.iter().map(|c| c.len() as u32).sum();

        // Write container header
        write_ppt_container_header(&mut container, record_type::EX_HYPERLINK, content_size)?;

        // Write children
        for child in children {
            container.extend_from_slice(&child);
        }

        Ok(container)
    }
}

// =============================================================================
// Shape Hyperlink Attachment
// =============================================================================

/// Hyperlink attached to a shape (for mouse click or hover)
#[derive(Debug, Clone)]
pub struct ShapeHyperlink {
    /// Hyperlink reference ID
    pub hyperlink_id: u32,
    /// Is this a mouse click action (vs hover)
    pub on_click: bool,
    /// Play sound (sound reference ID, 0 for none)
    pub sound_ref: u32,
    /// Highlight mode
    pub highlight: bool,
}

impl ShapeHyperlink {
    /// Create click hyperlink
    pub fn click(hyperlink_id: u32) -> Self {
        Self {
            hyperlink_id,
            on_click: true,
            sound_ref: 0,
            highlight: true,
        }
    }

    /// Create hover hyperlink
    pub fn hover(hyperlink_id: u32) -> Self {
        Self {
            hyperlink_id,
            on_click: false,
            sound_ref: 0,
            highlight: false,
        }
    }

    /// Set sound
    pub fn with_sound(mut self, sound_ref: u32) -> Self {
        self.sound_ref = sound_ref;
        self
    }

    /// Build InteractiveInfo container for ClientData (per POI InteractiveInfo)
    pub fn build_interactive_info(
        &self,
        atom: &InteractiveInfoAtom,
    ) -> Result<Vec<u8>, std::io::Error> {
        let mut container = Vec::new();

        // InteractiveInfoAtom (16 bytes)
        let mut atom_record = Vec::new();
        write_ppt_header(
            &mut atom_record,
            record_type::INTERACTIVE_INFO_ATOM,
            InteractiveInfoAtom::SIZE as u32,
        )?;
        atom_record.extend_from_slice(atom.as_bytes());

        let content_size = atom_record.len() as u32;

        // Container header - POI uses instance=0, version=0x0F
        write_ppt_container_header(&mut container, record_type::INTERACTIVE_INFO, content_size)?;

        container.extend_from_slice(&atom_record);

        Ok(container)
    }
}

// =============================================================================
// Helper Functions
// =============================================================================

/// Write a PPT record header (8 bytes)
fn write_ppt_header<W: std::io::Write>(
    writer: &mut W,
    rec_type: u16,
    length: u32,
) -> Result<(), std::io::Error> {
    // Version=0, Instance=0
    writer.write_all(&[0x00, 0x00])?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Write a PPT container header (8 bytes)
fn write_ppt_container_header<W: std::io::Write>(
    writer: &mut W,
    rec_type: u16,
    length: u32,
) -> Result<(), std::io::Error> {
    // Version=0x0F (container)
    writer.write_all(&[0x0F, 0x00])?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Write a PPT container header with custom instance
#[allow(dead_code)]
fn write_ppt_container_header_with_instance<W: std::io::Write>(
    writer: &mut W,
    version: u8,
    instance: u16,
    rec_type: u16,
    length: u32,
) -> Result<(), std::io::Error> {
    let ver_inst = (version as u16 & 0x0F) | ((instance & 0x0FFF) << 4);
    writer.write_all(&ver_inst.to_le_bytes())?;
    writer.write_all(&rec_type.to_le_bytes())?;
    writer.write_all(&length.to_le_bytes())?;
    Ok(())
}

/// Build a CString record (UTF-16LE string) with instance value
#[allow(dead_code)]
fn build_cstring(instance: u16, text: &str) -> Result<Vec<u8>, std::io::Error> {
    // Convert instance to options format (instance << 4)
    build_cstring_with_options((instance & 0x0FFF) << 4, text)
}

/// Build a CString record (UTF-16LE string) with raw options value
/// Per POI CString: first 2 bytes are options (version + instance << 4)
fn build_cstring_with_options(options: u16, text: &str) -> Result<Vec<u8>, std::io::Error> {
    let mut record = Vec::new();

    // Convert to UTF-16LE
    let utf16: Vec<u16> = text.encode_utf16().collect();
    let data_len = (utf16.len() * 2) as u32;

    // Header with options (version=0, instance in upper 12 bits)
    record.extend_from_slice(&options.to_le_bytes());
    record.extend_from_slice(&record_type::CSTRING.to_le_bytes());
    record.extend_from_slice(&data_len.to_le_bytes());

    // String data
    for ch in utf16 {
        record.extend_from_slice(&ch.to_le_bytes());
    }

    Ok(record)
}

// =============================================================================
// Tests
// =============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_hyperlink_creation() {
        let link = Hyperlink::url("https://example.com").with_display_text("Example");
        assert!(link.is_external());
        assert_eq!(link.target_string(), Some("https://example.com"));
    }

    #[test]
    fn test_hyperlink_collection() {
        let mut collection = HyperlinkCollection::new();
        let id1 = collection.add(Hyperlink::url("https://a.com"));
        let id2 = collection.add(Hyperlink::url("https://b.com"));

        assert_eq!(id1, 1);
        assert_eq!(id2, 2);
        assert_eq!(collection.len(), 2);
    }

    #[test]
    fn test_interactive_info_atom() {
        let atom = InteractiveInfoAtom::url_link(1);
        assert_eq!(atom.action, HyperlinkAction::Hyperlink as u8);
        // Copy from packed struct to avoid unaligned access
        let hyperlink_ref = { atom.hyperlink_ref };
        assert_eq!(hyperlink_ref, 1);
    }

    #[test]
    fn test_cstring_build() {
        let cstring = build_cstring(0, "Test").unwrap();
        // Header (8) + "Test" in UTF-16LE (8)
        assert_eq!(cstring.len(), 16);
    }
}
