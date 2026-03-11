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
    /// Maps (slide_index, media_index_in_slide) to (video_rel_id, media_rel_id, poster_rel_id)
    /// - video_rel_id: OOXML video/audio relationship (for r:link in videoFile/audioFile)
    /// - media_rel_id: Microsoft media relationship (for r:embed in p14:media)
    /// - poster_rel_id: Poster image relationship (for r:embed in blipFill/blip)
    media_ids: HashMap<(usize, usize), (String, String, String)>,
    /// Maps slide_index to comments relationship ID
    comments_ids: HashMap<usize, String>,
    /// Maps (slide_index, chart_idx) to chart relationship ID
    chart_ids: HashMap<(usize, u32), String>,
    /// Maps (slide_index, diagram_idx) to SmartArt relationship IDs
    /// (data_rel_id, layout_rel_id, style_rel_id, colors_rel_id)
    smartart_ids: HashMap<(usize, u32), (String, String, String, String)>,
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

    /// Add a media relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `media_index_in_slide` - The index of the media within that slide (0-based)
    /// * `video_rel_id` - The OOXML video/audio relationship ID (for r:link)
    /// * `media_rel_id` - The Microsoft media relationship ID (for r:embed in p14:media)
    /// * `poster_rel_id` - The poster image relationship ID (for r:embed in blipFill/blip)
    pub fn add_media(
        &mut self,
        slide_index: usize,
        media_index_in_slide: usize,
        video_rel_id: String,
        media_rel_id: String,
        poster_rel_id: String,
    ) {
        self.media_ids.insert(
            (slide_index, media_index_in_slide),
            (video_rel_id, media_rel_id, poster_rel_id),
        );
    }

    /// Get the relationship IDs for a media element in a specific slide.
    ///
    /// Returns (video_rel_id, media_rel_id, poster_rel_id) where:
    /// - video_rel_id is for r:link in a:videoFile/a:audioFile
    /// - media_rel_id is for r:embed in p14:media
    /// - poster_rel_id is for r:embed in blipFill/blip
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `media_index_in_slide` - The index of the media within that slide (0-based)
    pub fn get_media_ids(
        &self,
        slide_index: usize,
        media_index_in_slide: usize,
    ) -> Option<(&str, &str, &str)> {
        self.media_ids
            .get(&(slide_index, media_index_in_slide))
            .map(|(v, m, p)| (v.as_str(), m.as_str(), p.as_str()))
    }

    /// Add a comments relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `rel_id` - The relationship ID (e.g., "rId6")
    pub fn add_comments(&mut self, slide_index: usize, rel_id: String) {
        self.comments_ids.insert(slide_index, rel_id);
    }

    /// Get the comments relationship ID for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    #[allow(dead_code)] // Public API for future use
    pub fn get_comments_id(&self, slide_index: usize) -> Option<&str> {
        self.comments_ids.get(&slide_index).map(|s| s.as_str())
    }

    /// Add a chart relationship mapping for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `chart_idx` - The chart index (1-based, from presentation)
    /// * `rel_id` - The relationship ID (e.g., "rId7")
    pub fn add_chart(&mut self, slide_index: usize, chart_idx: u32, rel_id: String) {
        self.chart_ids.insert((slide_index, chart_idx), rel_id);
    }

    /// Get the chart relationship ID for a specific slide and chart.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `chart_idx` - The chart index (1-based)
    pub fn get_chart_id(&self, slide_index: usize, chart_idx: u32) -> Option<&str> {
        self.chart_ids
            .get(&(slide_index, chart_idx))
            .map(|s| s.as_str())
    }

    /// Add SmartArt relationship mappings for a specific slide.
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `diagram_idx` - The diagram index (1-based, from presentation)
    /// * `data_rel_id` - The data relationship ID
    /// * `layout_rel_id` - The layout relationship ID
    /// * `style_rel_id` - The style relationship ID
    /// * `colors_rel_id` - The colors relationship ID
    pub fn add_smartart(
        &mut self,
        slide_index: usize,
        diagram_idx: u32,
        data_rel_id: String,
        layout_rel_id: String,
        style_rel_id: String,
        colors_rel_id: String,
    ) {
        self.smartart_ids.insert(
            (slide_index, diagram_idx),
            (data_rel_id, layout_rel_id, style_rel_id, colors_rel_id),
        );
    }

    /// Get the SmartArt relationship IDs for a specific slide and diagram.
    ///
    /// Returns (data_rel_id, layout_rel_id, style_rel_id, colors_rel_id).
    ///
    /// # Arguments
    /// * `slide_index` - The index of the slide (0-based)
    /// * `diagram_idx` - The diagram index (1-based)
    pub fn get_smartart_ids(
        &self,
        slide_index: usize,
        diagram_idx: u32,
    ) -> Option<(&str, &str, &str, &str)> {
        self.smartart_ids
            .get(&(slide_index, diagram_idx))
            .map(|(d, l, s, c)| (d.as_str(), l.as_str(), s.as_str(), c.as_str()))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_relationship_mapper_new() {
        let mapper = RelationshipMapper::new();
        assert!(mapper.image_ids.is_empty());
        assert!(mapper.notes_ids.is_empty());
        assert!(mapper.background_ids.is_empty());
    }

    #[test]
    fn test_relationship_mapper_default() {
        let mapper: RelationshipMapper = Default::default();
        assert!(mapper.image_ids.is_empty());
    }

    #[test]
    fn test_add_and_get_image() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_image(0, 0, "rId2".to_string());
        mapper.add_image(0, 1, "rId3".to_string());
        mapper.add_image(1, 0, "rId4".to_string());

        assert_eq!(mapper.get_image_id(0, 0), Some("rId2"));
        assert_eq!(mapper.get_image_id(0, 1), Some("rId3"));
        assert_eq!(mapper.get_image_id(1, 0), Some("rId4"));
        assert_eq!(mapper.get_image_id(0, 2), None);
        assert_eq!(mapper.get_image_id(2, 0), None);
    }

    #[test]
    fn test_add_and_get_notes() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_notes(0, "rId5".to_string());
        mapper.add_notes(1, "rId6".to_string());

        assert_eq!(mapper.get_notes_id(0), Some("rId5"));
        assert_eq!(mapper.get_notes_id(1), Some("rId6"));
        assert_eq!(mapper.get_notes_id(2), None);
    }

    #[test]
    fn test_add_and_get_background() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_background(0, "rId7".to_string());

        assert_eq!(mapper.get_background_id(0), Some("rId7"));
        assert_eq!(mapper.get_background_id(1), None);
    }

    #[test]
    fn test_add_and_get_media() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_media(
            0,
            0,
            "rIdVideo1".to_string(),
            "rIdMedia1".to_string(),
            "rIdPoster1".to_string(),
        );
        mapper.add_media(
            0,
            1,
            "rIdVideo2".to_string(),
            "rIdMedia2".to_string(),
            "rIdPoster2".to_string(),
        );

        let result = mapper.get_media_ids(0, 0);
        assert!(result.is_some());
        let (video, media, poster) = result.unwrap();
        assert_eq!(video, "rIdVideo1");
        assert_eq!(media, "rIdMedia1");
        assert_eq!(poster, "rIdPoster1");

        let result2 = mapper.get_media_ids(0, 1);
        assert!(result2.is_some());
        let (video2, media2, poster2) = result2.unwrap();
        assert_eq!(video2, "rIdVideo2");
        assert_eq!(media2, "rIdMedia2");
        assert_eq!(poster2, "rIdPoster2");

        assert_eq!(mapper.get_media_ids(0, 2), None);
        assert_eq!(mapper.get_media_ids(1, 0), None);
    }

    #[test]
    fn test_add_and_get_comments() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_comments(0, "rId8".to_string());
        mapper.add_comments(2, "rId9".to_string());

        assert_eq!(mapper.get_comments_id(0), Some("rId8"));
        assert_eq!(mapper.get_comments_id(2), Some("rId9"));
        assert_eq!(mapper.get_comments_id(1), None);
    }

    #[test]
    fn test_add_and_get_chart() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_chart(0, 1, "rIdChart1".to_string());
        mapper.add_chart(0, 2, "rIdChart2".to_string());
        mapper.add_chart(1, 1, "rIdChart3".to_string());

        assert_eq!(mapper.get_chart_id(0, 1), Some("rIdChart1"));
        assert_eq!(mapper.get_chart_id(0, 2), Some("rIdChart2"));
        assert_eq!(mapper.get_chart_id(1, 1), Some("rIdChart3"));
        assert_eq!(mapper.get_chart_id(0, 3), None);
        assert_eq!(mapper.get_chart_id(2, 1), None);
    }

    #[test]
    fn test_add_and_get_smartart() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_smartart(
            0,
            1,
            "rIdData1".to_string(),
            "rIdLayout1".to_string(),
            "rIdStyle1".to_string(),
            "rIdColors1".to_string(),
        );

        let result = mapper.get_smartart_ids(0, 1);
        assert!(result.is_some());
        let (data, layout, style, colors) = result.unwrap();
        assert_eq!(data, "rIdData1");
        assert_eq!(layout, "rIdLayout1");
        assert_eq!(style, "rIdStyle1");
        assert_eq!(colors, "rIdColors1");

        assert_eq!(mapper.get_smartart_ids(0, 2), None);
        assert_eq!(mapper.get_smartart_ids(1, 1), None);
    }

    #[test]
    fn test_multiple_relationships_per_slide() {
        let mut mapper = RelationshipMapper::new();
        // Add multiple types of relationships for the same slide
        mapper.add_image(0, 0, "rIdImage1".to_string());
        mapper.add_notes(0, "rIdNotes1".to_string());
        mapper.add_background(0, "rIdBg1".to_string());
        mapper.add_chart(0, 1, "rIdChart1".to_string());

        assert_eq!(mapper.get_image_id(0, 0), Some("rIdImage1"));
        assert_eq!(mapper.get_notes_id(0), Some("rIdNotes1"));
        assert_eq!(mapper.get_background_id(0), Some("rIdBg1"));
        assert_eq!(mapper.get_chart_id(0, 1), Some("rIdChart1"));
    }

    #[test]
    fn test_overwrite_relationship() {
        let mut mapper = RelationshipMapper::new();
        mapper.add_image(0, 0, "rId1".to_string());
        assert_eq!(mapper.get_image_id(0, 0), Some("rId1"));

        // Overwrite with new relationship ID
        mapper.add_image(0, 0, "rId2".to_string());
        assert_eq!(mapper.get_image_id(0, 0), Some("rId2"));
    }
}
