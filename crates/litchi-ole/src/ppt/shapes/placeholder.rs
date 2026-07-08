/// Placeholder shape implementation.
///
/// Placeholders are special shapes that define the layout and structure
/// of PowerPoint slides. They represent positions where content like titles,
/// text, charts, or media should be placed.
use super::shape::{Shape, ShapeContainer, ShapeProperties};

/// Placeholder size options (quarter, half, full).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PlaceholderSize {
    /// Quarter size placeholder
    Quarter,
    /// Half size placeholder
    Half,
    /// Full size placeholder
    Full,
}

/// Types of placeholders in PowerPoint presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PlaceholderType {
    /// No placeholder shape
    None,
    /// Title placeholder
    Title,
    /// Body/content placeholder
    Body,
    /// Center title placeholder
    CenterTitle,
    /// Subtitle placeholder
    SubTitle,
    /// Chart placeholder
    Chart,
    /// Table placeholder
    Table,
    /// Clip art placeholder
    ClipArt,
    /// Diagram placeholder
    Diagram,
    /// Media clip placeholder
    MediaClip,
    /// Object placeholder (embedded objects)
    Object,
    /// Content placeholder
    Content,
    /// Picture placeholder
    Picture,
    /// Slide image placeholder
    SlideImage,
    /// Vertical text placeholder
    VerticalTextTitle,
    /// Vertical text body placeholder
    VerticalTextBody,
    /// Notes slide image placeholder
    NotesSlideImage,
    /// Notes slide text placeholder
    NotesSlideText,
    /// Header placeholder
    Header,
    /// Footer placeholder
    Footer,
    /// Slide number placeholder
    SlideNumber,
    /// Date and time placeholder
    DateAndTime,
    /// Vertical object placeholder
    VerticalObject,
    /// Copyright placeholder
    Copyright,
    /// Custom placeholder
    Custom(u16),
}

impl From<u16> for PlaceholderType {
    fn from(value: u16) -> Self {
        match value {
            0 => PlaceholderType::None,
            1 => PlaceholderType::Title,
            2 => PlaceholderType::Body,
            3 => PlaceholderType::CenterTitle,
            4 => PlaceholderType::SubTitle,
            5 => PlaceholderType::Chart,
            6 => PlaceholderType::Table,
            7 => PlaceholderType::ClipArt,
            8 => PlaceholderType::Diagram,
            9 => PlaceholderType::MediaClip,
            10 => PlaceholderType::Object,
            11 => PlaceholderType::SlideImage,
            19 => PlaceholderType::Content,
            26 => PlaceholderType::Picture,
            12 => PlaceholderType::VerticalTextTitle,
            13 => PlaceholderType::VerticalTextBody,
            14 => PlaceholderType::NotesSlideImage,
            15 => PlaceholderType::NotesSlideText,
            16 => PlaceholderType::Header,
            17 => PlaceholderType::Footer,
            18 => PlaceholderType::SlideNumber,
            20 => PlaceholderType::VerticalObject,
            21 => PlaceholderType::Copyright,
            other => PlaceholderType::Custom(other),
        }
    }
}

impl std::fmt::Display for PlaceholderType {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PlaceholderType::None => write!(f, "None"),
            PlaceholderType::Title => write!(f, "Title"),
            PlaceholderType::Body => write!(f, "Body"),
            PlaceholderType::CenterTitle => write!(f, "CenterTitle"),
            PlaceholderType::SubTitle => write!(f, "SubTitle"),
            PlaceholderType::Chart => write!(f, "Chart"),
            PlaceholderType::Table => write!(f, "Table"),
            PlaceholderType::ClipArt => write!(f, "ClipArt"),
            PlaceholderType::Diagram => write!(f, "Diagram"),
            PlaceholderType::MediaClip => write!(f, "MediaClip"),
            PlaceholderType::Object => write!(f, "Object"),
            PlaceholderType::Content => write!(f, "Content"),
            PlaceholderType::Picture => write!(f, "Picture"),
            PlaceholderType::SlideImage => write!(f, "SlideImage"),
            PlaceholderType::VerticalTextTitle => write!(f, "VerticalTextTitle"),
            PlaceholderType::VerticalTextBody => write!(f, "VerticalTextBody"),
            PlaceholderType::NotesSlideImage => write!(f, "NotesSlideImage"),
            PlaceholderType::NotesSlideText => write!(f, "NotesSlideText"),
            PlaceholderType::Header => write!(f, "Header"),
            PlaceholderType::Footer => write!(f, "Footer"),
            PlaceholderType::SlideNumber => write!(f, "SlideNumber"),
            PlaceholderType::DateAndTime => write!(f, "DateAndTime"),
            PlaceholderType::VerticalObject => write!(f, "VerticalObject"),
            PlaceholderType::Copyright => write!(f, "Copyright"),
            PlaceholderType::Custom(id) => write!(f, "Custom({})", id),
        }
    }
}

/// A placeholder shape in a PowerPoint presentation.
///
/// Uses lifetime parameter `'a` to enable zero-copy parsing when the shape
/// data can be borrowed from a larger buffer.
#[derive(Debug, Clone)]
pub struct Placeholder<'a> {
    /// Shape container with properties and data
    container: ShapeContainer<'a>,
    /// The type of placeholder
    placeholder_type: PlaceholderType,
    /// Placeholder size (index in the layout)
    size: Option<u8>,
    /// Placeholder index within the slide
    index: Option<u16>,
    /// Raw placeholder data for advanced parsing
    raw_placeholder_data: Option<Vec<u8>>,
}

impl<'a> Placeholder<'a> {
    /// Create a new placeholder shape with owned data.
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            placeholder_type: PlaceholderType::Body, // Default
            size: None,
            index: None,
            raw_placeholder_data: None,
        }
    }

    /// Create a placeholder from an Escher record with zero-copy parsing.
    pub fn from_escher_record(
        record: &'a super::escher::EscherRecord<'a>,
    ) -> super::super::package::Result<Self> {
        // Extract basic shape properties
        let properties = record.extract_shape_properties()?;

        // Extract placeholder information from OEPlaceholderAtom or similar records
        let placeholder_info = Self::extract_placeholder_info_from_record(record)?;

        let (placeholder_type, size, index) =
            if let Some((poi_id, size_val, index_val)) = placeholder_info {
                match Self::map_poi_placeholder_id_to_type(poi_id) {
                    Ok(placeholder_type) => (placeholder_type, Some(size_val), Some(index_val)),
                    Err(_) => (PlaceholderType::Body, None, None), // Fallback
                }
            } else {
                (PlaceholderType::Body, None, None)
            };

        // Extract additional placeholder properties from Escher data with zero-copy
        let container = ShapeContainer::new_borrowed(properties, &record.data);

        Ok(Self {
            container,
            placeholder_type,
            size,
            index,
            raw_placeholder_data: Some(record.data.to_vec()),
        })
    }

    /// Create a placeholder from an existing container.
    ///
    /// Based on POI's HSLFPlaceholder which extracts placeholder info from
    /// the shape's client data records (OEPlaceholderAtom).
    pub fn from_container(container: ShapeContainer<'a>) -> Self {
        // Try to extract placeholder information from the raw data
        // The raw data may contain an OEPlaceholderAtom record
        let (placeholder_type, size, index) = if !container.raw_data.is_empty() {
            // Parse as Escher record to find OEPlaceholderAtom
            if let Ok((escher_record, _)) =
                super::escher::EscherRecord::parse(&container.raw_data, 0)
            {
                if let Ok(Some((poi_id, size_val, index_val))) =
                    escher_record.extract_placeholder_info()
                {
                    // Map POI placeholder ID to our type
                    let ph_type = Self::map_poi_placeholder_id_to_type(poi_id)
                        .unwrap_or(PlaceholderType::Body);
                    (ph_type, Some(size_val), Some(index_val))
                } else {
                    (PlaceholderType::Body, None, None)
                }
            } else {
                (PlaceholderType::Body, None, None)
            }
        } else {
            (PlaceholderType::Body, None, None)
        };

        Self {
            container,
            placeholder_type,
            size,
            index,
            raw_placeholder_data: None,
        }
    }

    /// Extract placeholder information from an Escher record.
    /// This follows POI's logic for parsing OEPlaceholderAtom records.
    fn extract_placeholder_info_from_record(
        record: &super::escher::EscherRecord,
    ) -> super::super::package::Result<Option<(u16, u8, u16)>> {
        // Use the EscherRecord's extract_placeholder_info method
        record.extract_placeholder_info()
    }

    /// Map POI placeholder IDs to our placeholder types.
    /// This follows the mapping used in POI's Placeholder enum.
    fn map_poi_placeholder_id_to_type(
        poi_id: u16,
    ) -> super::super::package::Result<PlaceholderType> {
        // Based on POI's Placeholder enum mapping
        let placeholder_type = match poi_id {
            0 => PlaceholderType::None,
            13 => PlaceholderType::Title,             // TITLE
            14 => PlaceholderType::Body,              // BODY
            15 => PlaceholderType::CenterTitle,       // CENTERED_TITLE
            16 => PlaceholderType::SubTitle,          // SUBTITLE
            7 => PlaceholderType::DateAndTime,        // DATETIME
            8 => PlaceholderType::SlideNumber,        // SLIDE_NUMBER
            9 => PlaceholderType::Footer,             // FOOTER
            10 => PlaceholderType::Header,            // HEADER
            19 => PlaceholderType::Content,           // CONTENT
            20 => PlaceholderType::Chart,             // CHART
            21 => PlaceholderType::Table,             // TABLE
            22 => PlaceholderType::ClipArt,           // CLIP_ART
            23 => PlaceholderType::Diagram,           // DGM
            24 => PlaceholderType::MediaClip,         // MEDIA
            11 => PlaceholderType::SlideImage,        // SLIDE_IMAGE
            26 => PlaceholderType::Picture,           // PICTURE
            25 => PlaceholderType::VerticalObject,    // VERTICAL_OBJECT
            17 => PlaceholderType::VerticalTextTitle, // VERTICAL_TEXT_TITLE
            18 => PlaceholderType::VerticalTextBody,  // VERTICAL_TEXT_BODY
            _ => {
                // For unknown types, return None (POI doesn't have this)
                return Err(super::super::package::PptError::Corrupted(format!(
                    "Unknown placeholder type ID: {}",
                    poi_id
                )));
            },
        };

        Ok(placeholder_type)
    }

    /// Get the raw placeholder data for advanced parsing.
    pub fn raw_placeholder_data(&self) -> Option<&[u8]> {
        self.raw_placeholder_data.as_deref()
    }

    /// Get the placeholder type.
    pub fn placeholder_type(&self) -> PlaceholderType {
        self.placeholder_type
    }

    /// Set the placeholder type.
    pub fn set_placeholder_type(&mut self, placeholder_type: PlaceholderType) {
        self.placeholder_type = placeholder_type;
    }

    /// Get the placeholder size (layout index).
    pub fn size(&self) -> Option<u8> {
        self.size
    }

    /// Set the placeholder size.
    pub fn set_size(&mut self, size: u8) {
        self.size = Some(size);
    }

    /// Get the placeholder index within the slide.
    pub fn index(&self) -> Option<u16> {
        self.index
    }

    /// Set the placeholder index.
    pub fn set_index(&mut self, index: u16) {
        self.index = Some(index);
    }

    /// Check if this is a title placeholder.
    pub fn is_title(&self) -> bool {
        matches!(
            self.placeholder_type,
            PlaceholderType::Title | PlaceholderType::CenterTitle | PlaceholderType::SubTitle
        )
    }

    /// Check if this is a content/body placeholder.
    pub fn is_content(&self) -> bool {
        matches!(self.placeholder_type, PlaceholderType::Body)
    }

    /// Check if this is a media placeholder (picture, chart, etc.).
    pub fn is_media(&self) -> bool {
        matches!(
            self.placeholder_type,
            PlaceholderType::Picture
                | PlaceholderType::Chart
                | PlaceholderType::Table
                | PlaceholderType::ClipArt
                | PlaceholderType::Diagram
                | PlaceholderType::MediaClip
                | PlaceholderType::Object
                | PlaceholderType::Content
                | PlaceholderType::SlideImage
        )
    }

    /// Get the placeholder size (quarter, half, full).
    pub fn placeholder_size(&self) -> PlaceholderSize {
        match self.size {
            Some(1) => PlaceholderSize::Quarter,
            Some(2) => PlaceholderSize::Half,
            Some(3) => PlaceholderSize::Full,
            _ => PlaceholderSize::Full, // Default to full size
        }
    }

    /// Set the placeholder size.
    pub fn set_placeholder_size(&mut self, size: PlaceholderSize) {
        self.size = match size {
            PlaceholderSize::Quarter => Some(1),
            PlaceholderSize::Half => Some(2),
            PlaceholderSize::Full => Some(3),
        };
    }
}

impl<'a> Shape for Placeholder<'a>
where
    'a: 'static,
{
    fn properties(&self) -> &ShapeProperties {
        &self.container.properties
    }

    fn properties_mut(&mut self) -> &mut ShapeProperties {
        &mut self.container.properties
    }

    fn text(&self) -> super::super::package::Result<String> {
        // Placeholders typically don't have their own text - they hold content
        Ok(String::new())
    }

    fn has_text(&self) -> bool {
        false // Placeholders don't have inherent text content
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::super::shape::ShapeType;
    use super::*;

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_placeholder_creation() {
        let mut props = ShapeProperties::default();
        props.id = 2001;
        props.shape_type = ShapeType::Placeholder;
        props.x = 50;
        props.y = 50;
        props.width = 400;
        props.height = 300;

        let placeholder = Placeholder::new(props, vec![1, 2, 3]);
        assert_eq!(placeholder.id(), 2001);
        assert_eq!(placeholder.shape_type(), ShapeType::Placeholder);
        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Body);
        assert!(placeholder.is_content());
        assert!(!placeholder.is_title());
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_placeholder_type_operations() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::Placeholder;

        let mut placeholder = Placeholder::new(props, vec![]);
        placeholder.set_placeholder_type(PlaceholderType::Title);
        placeholder.set_size(1);
        placeholder.set_index(0);

        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
        assert_eq!(placeholder.size(), Some(1));
        assert_eq!(placeholder.index(), Some(0));
        assert!(placeholder.is_title());
        assert!(!placeholder.is_content());
    }

    #[test]
    #[allow(clippy::field_reassign_with_default)]
    fn test_placeholder_media_check() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::Placeholder;

        let mut placeholder = Placeholder::new(props, vec![]);

        // Test picture placeholder
        placeholder.set_placeholder_type(PlaceholderType::Picture);
        assert!(placeholder.is_media());
        assert!(!placeholder.is_title());

        // Test title placeholder
        placeholder.set_placeholder_type(PlaceholderType::Title);
        assert!(!placeholder.is_media());
        assert!(placeholder.is_title());
    }

    #[test]
    fn test_placeholder_type_conversion() {
        assert_eq!(PlaceholderType::from(1), PlaceholderType::Title);
        assert_eq!(PlaceholderType::from(2), PlaceholderType::Body);
        assert_eq!(PlaceholderType::from(11), PlaceholderType::SlideImage);
        assert_eq!(PlaceholderType::from(26), PlaceholderType::Picture);
        assert_eq!(PlaceholderType::from(999), PlaceholderType::Custom(999));
    }
}
