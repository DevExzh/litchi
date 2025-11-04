/// Relationship ID mapper for tracking relationships during presentation generation.
///
/// This module handles the mapping between presentation elements (images, notes)
/// and their relationship IDs in the OOXML package. Unlike the placeholder approach,
/// this creates relationships first and provides the actual IDs for XML generation.
use std::collections::HashMap;

/// Maps presentation elements to their relationship IDs.
///
/// This is used during presentation generation to track which relationship ID
/// corresponds to which image or notes slide, allowing proper XML generation
/// with actual relationship references.
///
/// The mapper is organized per-slide, as each slide has its own set of relationships.
#[derive(Debug, Default)]
pub struct RelationshipMapper {
    /// Maps (slide_index, image_index_in_slide) to relationship ID
    image_ids: HashMap<(usize, usize), String>,
    /// Maps slide_index to notes slide relationship ID
    notes_ids: HashMap<usize, String>,
    /// Maps slide_index to background image relationship ID
    background_ids: HashMap<usize, String>,
}

impl RelationshipMapper {
    /// Create a new empty relationship mapper.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add an image relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `image_index_in_slide` - The index of the image within that slide (0-based)
    /// * `rel_id` - The relationship ID (e.g., "rId2")
    pub fn add_image(&mut self, slide_index: usize, image_index_in_slide: usize, rel_id: String) {
        self.image_ids
            .insert((slide_index, image_index_in_slide), rel_id);
    }

    /// Add a notes slide relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `rel_id` - The relationship ID (e.g., "rId3")
    pub fn add_notes(&mut self, slide_index: usize, rel_id: String) {
        self.notes_ids.insert(slide_index, rel_id);
    }

    /// Get the relationship ID for an image in a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `image_index_in_slide` - The index of the image within that slide (0-based)
    pub fn get_image_id(&self, slide_index: usize, image_index_in_slide: usize) -> Option<&str> {
        self.image_ids
            .get(&(slide_index, image_index_in_slide))
            .map(|s| s.as_str())
    }

    /// Get the notes slide relationship ID for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    #[allow(dead_code)] // Public API but not used in the current implementation
    pub fn get_notes_id(&self, slide_index: usize) -> Option<&str> {
        self.notes_ids.get(&slide_index).map(|s| s.as_str())
    }

    /// Add a background image relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `rel_id` - The relationship ID (e.g., "rId4")
    pub fn add_background(&mut self, slide_index: usize, rel_id: String) {
        self.background_ids.insert(slide_index, rel_id);
    }

    /// Get the background image relationship ID for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    pub fn get_background_id(&self, slide_index: usize) -> Option<&str> {
        self.background_ids.get(&slide_index).map(|s| s.as_str())
    }
}
