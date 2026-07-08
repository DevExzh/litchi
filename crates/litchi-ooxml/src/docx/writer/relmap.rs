/// Relationship ID mapper for tracking relationships during document generation.
///
/// This module handles the mapping between document elements (hyperlinks, images)
/// and their relationship IDs in the OOXML package. Unlike the placeholder approach,
/// this creates relationships first and provides the actual IDs for XML generation.
use std::collections::HashMap;

/// Maps document elements to their relationship IDs.
///
/// This is used during document generation to track which relationship ID
/// corresponds to which hyperlink or image, allowing proper XML generation
/// with actual relationship references.
#[derive(Debug, Default)]
pub struct RelationshipMapper {
    /// Maps hyperlink index to relationship ID
    hyperlink_ids: HashMap<usize, String>,
    /// Maps image index to relationship ID
    image_ids: HashMap<usize, String>,
    /// Header relationship ID (if any)
    header_id: Option<String>,
    /// Footer relationship ID (if any)
    footer_id: Option<String>,
    /// Footnotes relationship ID (if any)
    footnotes_id: Option<String>,
    /// Endnotes relationship ID (if any)
    endnotes_id: Option<String>,
}

impl RelationshipMapper {
    /// Create a new empty relationship mapper.
    pub fn new() -> Self {
        Self::default()
    }

    /// Add a hyperlink relationship mapping.
    pub fn add_hyperlink(&mut self, index: usize, rel_id: String) {
        self.hyperlink_ids.insert(index, rel_id);
    }

    /// Add an image relationship mapping.
    pub fn add_image(&mut self, index: usize, rel_id: String) {
        self.image_ids.insert(index, rel_id);
    }

    /// Set the header relationship ID.
    pub fn set_header_id(&mut self, rel_id: String) {
        self.header_id = Some(rel_id);
    }

    /// Set the footer relationship ID.
    pub fn set_footer_id(&mut self, rel_id: String) {
        self.footer_id = Some(rel_id);
    }

    /// Get the relationship ID for a hyperlink by index.
    pub fn get_hyperlink_id(&self, index: usize) -> Option<&str> {
        self.hyperlink_ids.get(&index).map(|s| s.as_str())
    }

    /// Get the relationship ID for an image by index.
    pub fn get_image_id(&self, index: usize) -> Option<&str> {
        self.image_ids.get(&index).map(|s| s.as_str())
    }

    /// Get the header relationship ID.
    pub fn get_header_id(&self) -> Option<&str> {
        self.header_id.as_deref()
    }

    /// Get the footer relationship ID.
    pub fn get_footer_id(&self) -> Option<&str> {
        self.footer_id.as_deref()
    }

    /// Set the footnotes relationship ID.
    pub fn set_footnotes_id(&mut self, rel_id: String) {
        self.footnotes_id = Some(rel_id);
    }

    /// Get the footnotes relationship ID.
    pub fn get_footnotes_id(&self) -> Option<&str> {
        self.footnotes_id.as_deref()
    }

    /// Set the endnotes relationship ID.
    pub fn set_endnotes_id(&mut self, rel_id: String) {
        self.endnotes_id = Some(rel_id);
    }

    /// Get the endnotes relationship ID.
    pub fn get_endnotes_id(&self) -> Option<&str> {
        self.endnotes_id.as_deref()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_relationship_mapper_new() {
        let mapper = RelationshipMapper::new();
        assert!(mapper.hyperlink_ids.is_empty());
        assert!(mapper.image_ids.is_empty());
        assert!(mapper.header_id.is_none());
        assert!(mapper.footer_id.is_none());
        assert!(mapper.footnotes_id.is_none());
        assert!(mapper.endnotes_id.is_none());
    }

    #[test]
    fn test_relationship_mapper_default() {
        let mapper: RelationshipMapper = Default::default();
        assert!(mapper.hyperlink_ids.is_empty());
        assert!(mapper.image_ids.is_empty());
    }

    #[test]
    fn test_add_and_get_hyperlink() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_hyperlink(0, "rId1".to_string());
        mapper.add_hyperlink(1, "rId2".to_string());

        assert_eq!(mapper.get_hyperlink_id(0), Some("rId1"));
        assert_eq!(mapper.get_hyperlink_id(1), Some("rId2"));
        assert_eq!(mapper.get_hyperlink_id(2), None);
    }

    #[test]
    fn test_add_and_get_image() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_image(0, "rId5".to_string());
        mapper.add_image(1, "rId6".to_string());

        assert_eq!(mapper.get_image_id(0), Some("rId5"));
        assert_eq!(mapper.get_image_id(1), Some("rId6"));
        assert_eq!(mapper.get_image_id(99), None);
    }

    #[test]
    fn test_set_and_get_header_id() {
        let mut mapper = RelationshipMapper::new();
        assert_eq!(mapper.get_header_id(), None);

        mapper.set_header_id("rId10".to_string());
        assert_eq!(mapper.get_header_id(), Some("rId10"));
    }

    #[test]
    fn test_set_and_get_footer_id() {
        let mut mapper = RelationshipMapper::new();
        assert_eq!(mapper.get_footer_id(), None);

        mapper.set_footer_id("rId11".to_string());
        assert_eq!(mapper.get_footer_id(), Some("rId11"));
    }

    #[test]
    fn test_set_and_get_footnotes_id() {
        let mut mapper = RelationshipMapper::new();
        assert_eq!(mapper.get_footnotes_id(), None);

        mapper.set_footnotes_id("rId20".to_string());
        assert_eq!(mapper.get_footnotes_id(), Some("rId20"));
    }

    #[test]
    fn test_set_and_get_endnotes_id() {
        let mut mapper = RelationshipMapper::new();
        assert_eq!(mapper.get_endnotes_id(), None);

        mapper.set_endnotes_id("rId21".to_string());
        assert_eq!(mapper.get_endnotes_id(), Some("rId21"));
    }

    #[test]
    fn test_relationship_mapper_debug() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_hyperlink(0, "rId1".to_string());
        let debug_str = format!("{:?}", mapper);
        assert!(debug_str.contains("RelationshipMapper"));
    }
}
