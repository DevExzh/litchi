//! Custom slide show (named show) support for PPT files.
//!
//! Implements NamedShows/NamedShow/NamedShowSlides records per [MS-PPT].
//! Custom shows are stored inside the DocumentContainer.
//!
//! # Binary Structure
//!
//! ```text
//! NamedShows (container, type=1040)
//! └── NamedShow (container, type=1041, instance=index)
//!     ├── CString (type=4026): show name (UTF-16LE, max 31 chars)
//!     └── NamedShowSlides (atom, type=1042)
//!         └── array of u32 slide IDs (0x100 + slide_index)
//! ```

use super::records::{PptError, RecordBuilder, record_type};

/// A custom (named) slide show definition.
///
/// Custom shows allow defining named subsets of slides that can be
/// presented independently from the main slide deck.
#[derive(Debug, Clone)]
pub struct CustomShow {
    /// Show name (max 31 characters, truncated if longer).
    pub name: String,
    /// Slide indices (0-based) included in this show, in presentation order.
    pub slide_indices: Vec<usize>,
}

impl CustomShow {
    /// Create a new custom show with the given name and slide indices.
    ///
    /// # Arguments
    ///
    /// * `name` - Show name (max 31 characters)
    /// * `slide_indices` - 0-based slide indices in presentation order
    ///
    /// # Example
    ///
    /// ```
    /// use litchi::ole::ppt::writer::custom_shows::CustomShow;
    /// // Create a custom show with slides 0, 2, and 4
    /// let show = CustomShow::new("Executive Summary", &[0, 2, 4]);
    /// ```
    pub fn new(name: &str, slide_indices: &[usize]) -> Self {
        Self {
            name: name.to_string(),
            slide_indices: slide_indices.to_vec(),
        }
    }
}

/// Build the NamedShows container for the Document.
///
/// Returns the serialized NamedShows container bytes, or an empty Vec if
/// there are no custom shows.
///
/// # Arguments
///
/// * `shows` - Slice of custom show definitions
pub fn build_named_shows(shows: &[CustomShow]) -> Result<Vec<u8>, PptError> {
    if shows.is_empty() {
        return Ok(Vec::new());
    }

    let mut children = Vec::new();

    for (i, show) in shows.iter().enumerate() {
        children.extend(build_named_show(show, i as u16)?);
    }

    // NamedShows container (type=1040)
    let mut container = RecordBuilder::new(0x0F, 0, record_type::NAMED_SHOWS);
    container.write_data(&children);
    container.build()
}

/// Build a single NamedShow container.
fn build_named_show(show: &CustomShow, index: u16) -> Result<Vec<u8>, PptError> {
    let mut children = Vec::new();

    // CString: show name (UTF-16LE, max 31 chars per LibreOffice)
    let name: String = show.name.chars().take(31).collect();
    let utf16: Vec<u16> = name.encode_utf16().collect();
    let mut name_data = Vec::with_capacity(utf16.len() * 2);
    for ch in &utf16 {
        name_data.extend_from_slice(&ch.to_le_bytes());
    }
    let mut cstr = RecordBuilder::new(0x00, 0, record_type::CSTRING);
    cstr.write_data(&name_data);
    children.extend(cstr.build()?);

    // NamedShowSlides atom: array of slide IDs
    // Per LibreOffice: slide ID = slide_index + 0x100
    let mut slides_data = Vec::with_capacity(show.slide_indices.len() * 4);
    for &idx in &show.slide_indices {
        let slide_id = (idx as u32) + 0x100;
        slides_data.extend_from_slice(&slide_id.to_le_bytes());
    }
    let mut slides_atom = RecordBuilder::new(0x00, 0, record_type::NAMED_SHOW_SLIDES);
    slides_atom.write_data(&slides_data);
    children.extend(slides_atom.build()?);

    // NamedShow container (type=1041, instance=show_index)
    let mut container = RecordBuilder::new(0x0F, index, record_type::NAMED_SHOW);
    container.write_data(&children);
    container.build()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_build_named_shows_empty() {
        let data = build_named_shows(&[]).unwrap();
        assert!(data.is_empty());
    }

    #[test]
    fn test_build_named_shows_single() {
        let shows = vec![CustomShow::new("Test Show", &[0, 1, 2])];
        let data = build_named_shows(&shows).unwrap();
        assert!(!data.is_empty());

        // Verify outer container type = 1040 (NamedShows)
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1040);
    }

    #[test]
    fn test_build_named_shows_multiple() {
        let shows = vec![
            CustomShow::new("Short Version", &[0, 3]),
            CustomShow::new("Full Version", &[0, 1, 2, 3, 4]),
        ];
        let data = build_named_shows(&shows).unwrap();
        assert!(!data.is_empty());

        // Verify outer container type = 1040 (NamedShows)
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1040);
    }

    #[test]
    fn test_slide_id_encoding() {
        let show = CustomShow::new("Test", &[0, 5, 10]);
        let data = build_named_show(&show, 0).unwrap();
        assert!(!data.is_empty());

        // Verify NamedShow container type = 1041
        let rtype = u16::from_le_bytes([data[2], data[3]]);
        assert_eq!(rtype, 1041);
    }
}
