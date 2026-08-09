/// Placeholder shape implementation.
///
/// Placeholders are special shapes that define the layout and structure
/// of `PowerPoint` slides. They represent positions where content like titles,
/// text, charts, or media should be placed.
use super::shape::{Shape, ShapeContainer, ShapeProperties};
use crate::odraw::ShapeExt as _;

/// Placeholder size options (quarter, half, full).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PlaceholderSize` is the established public API name; renaming it would break downstream crates"
)]
pub enum PlaceholderSize {
    /// Quarter size placeholder
    Quarter,
    /// Half size placeholder
    Half,
    /// Full size placeholder
    Full,
}

impl From<crate::AtomPlaceholderSize> for PlaceholderSize {
    fn from(size: crate::AtomPlaceholderSize) -> Self {
        match size {
            crate::AtomPlaceholderSize::Full => Self::Full,
            crate::AtomPlaceholderSize::Half => Self::Half,
            crate::AtomPlaceholderSize::Quarter => Self::Quarter,
        }
    }
}

/// Types of placeholders in `PowerPoint` presentations.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`PlaceholderType` is the established public API name; renaming it would break downstream crates"
)]
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

impl From<crate::PlaceholderKind> for PlaceholderType {
    fn from(kind: crate::PlaceholderKind) -> Self {
        use crate::PlaceholderKind as Kind;

        match kind {
            Kind::MasterTitle | Kind::Title => Self::Title,
            Kind::MasterBody | Kind::Body => Self::Body,
            Kind::MasterCenterTitle | Kind::CenterTitle => Self::CenterTitle,
            Kind::MasterSubTitle | Kind::SubTitle => Self::SubTitle,
            Kind::MasterNotesSlideImage | Kind::NotesSlideImage => Self::NotesSlideImage,
            Kind::MasterNotesBody | Kind::NotesBody => Self::NotesSlideText,
            Kind::MasterDate => Self::DateAndTime,
            Kind::MasterSlideNumber => Self::SlideNumber,
            Kind::MasterFooter => Self::Footer,
            Kind::MasterHeader => Self::Header,
            Kind::VerticalTitle => Self::VerticalTextTitle,
            Kind::VerticalBody => Self::VerticalTextBody,
            Kind::Object => Self::Object,
            Kind::Graph => Self::Chart,
            Kind::Table => Self::Table,
            Kind::ClipArt => Self::ClipArt,
            Kind::OrgChart => Self::Diagram,
            Kind::Media => Self::MediaClip,
            Kind::VerticalObject => Self::VerticalObject,
            Kind::Picture => Self::Picture,
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
            PlaceholderType::Custom(id) => write!(f, "Custom({id})"),
        }
    }
}

/// A placeholder shape in a `PowerPoint` presentation.
///
/// Uses lifetime parameter `'a` to enable zero-copy parsing when the shape
/// data can be borrowed from a larger buffer.
#[derive(Debug, Clone)]
pub struct Placeholder<'a> {
    /// Shape container with properties and data
    container: ShapeContainer<'a>,
    /// The type of placeholder
    kind: PlaceholderType,
    /// Placeholder size (index in the layout)
    size: Option<u8>,
    /// Placeholder index within the slide
    index: Option<u16>,
    /// Raw placeholder data for advanced parsing
    raw_placeholder_data: Option<Vec<u8>>,
}

impl<'a> Placeholder<'a> {
    /// Create a new placeholder shape with owned data.
    #[must_use]
    pub fn new(properties: ShapeProperties, raw_data: Vec<u8>) -> Self {
        Self {
            container: ShapeContainer::new(properties, raw_data),
            kind: PlaceholderType::Body, // Default
            size: None,
            index: None,
            raw_placeholder_data: None,
        }
    }

    pub(crate) fn from_parsed(
        properties: ShapeProperties,
        placeholder_type: PlaceholderType,
        size: PlaceholderSize,
        index: Option<u16>,
        text: Option<String>,
    ) -> Self {
        let mut container = ShapeContainer::new(properties, Vec::new());
        container.text_content = text;
        Self {
            container,
            kind: placeholder_type,
            size: Some(match size {
                PlaceholderSize::Quarter => 2,
                PlaceholderSize::Half => 1,
                PlaceholderSize::Full => 0,
            }),
            index,
            raw_placeholder_data: None,
        }
    }

    /// Create a placeholder from an existing container.
    ///
    /// Based on POI's `HSLFPlaceholder` which extracts placeholder info from
    /// the shape's client data records (`OEPlaceholderAtom`).
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_container(container: ShapeContainer<'a>) -> super::super::package::Result<Self> {
        let (placeholder_type, size, index) = if container.raw_data.is_empty() {
            (PlaceholderType::Body, None, None)
        } else {
            let (record, consumed) = litchi_odraw::Record::parse(&container.raw_data, 0)?;
            if consumed != container.raw_data.len() {
                return Err(super::super::package::Error::Corrupted(
                    "placeholder container has trailing OfficeArt records".to_owned(),
                ));
            }
            let shape = litchi_odraw::shape::Shape::try_from(record)?;
            match shape.placeholder()? {
                Some(info) => (
                    PlaceholderType::from(info.kind),
                    Some(match info.size {
                        crate::AtomPlaceholderSize::Full => 0,
                        crate::AtomPlaceholderSize::Half => 1,
                        crate::AtomPlaceholderSize::Quarter => 2,
                    }),
                    info.position,
                ),
                None => (PlaceholderType::Body, None, None),
            }
        };

        Ok(Self {
            container,
            kind: placeholder_type,
            size,
            index,
            raw_placeholder_data: None,
        })
    }

    /// Get the raw placeholder data for advanced parsing.
    #[must_use]
    pub fn raw_placeholder_data(&self) -> Option<&[u8]> {
        self.raw_placeholder_data.as_deref()
    }

    /// Get the placeholder type.
    #[must_use]
    pub fn placeholder_type(&self) -> PlaceholderType {
        self.kind
    }

    /// Set the placeholder type.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_placeholder_type(
        &mut self,
        placeholder_type: PlaceholderType,
    ) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Placeholder)?;
        self.kind = placeholder_type;
        Ok(())
    }

    /// Get the placeholder size (layout index).
    #[must_use]
    pub fn size(&self) -> Option<u8> {
        self.size
    }

    /// Set the placeholder size.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_size(&mut self, size: u8) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Placeholder)?;
        self.size = Some(size);
        Ok(())
    }

    /// Get the placeholder index within the slide.
    #[must_use]
    pub fn index(&self) -> Option<u16> {
        self.index
    }

    /// Set the placeholder index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_index(&mut self, index: u16) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Placeholder)?;
        self.index = Some(index);
        Ok(())
    }

    /// Check if this is a title placeholder.
    #[must_use]
    pub fn is_title(&self) -> bool {
        matches!(
            self.kind,
            PlaceholderType::Title | PlaceholderType::CenterTitle | PlaceholderType::SubTitle
        )
    }

    /// Check if this is a content/body placeholder.
    #[must_use]
    pub fn is_content(&self) -> bool {
        matches!(self.kind, PlaceholderType::Body | PlaceholderType::Content)
    }

    /// Check if this is a media placeholder (picture, chart, etc.).
    #[must_use]
    pub fn is_media(&self) -> bool {
        matches!(
            self.kind,
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
    #[must_use]
    pub fn placeholder_size(&self) -> PlaceholderSize {
        match self.size {
            Some(2) => PlaceholderSize::Quarter,
            Some(1) => PlaceholderSize::Half,
            _ => PlaceholderSize::Full,
        }
    }

    /// Set the placeholder size.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_placeholder_size(
        &mut self,
        size: PlaceholderSize,
    ) -> Result<(), super::shape::MutationError> {
        self.container
            .ensure_mutable(super::shape::Mutation::Placeholder)?;
        self.size = match size {
            PlaceholderSize::Quarter => Some(2),
            PlaceholderSize::Half => Some(1),
            PlaceholderSize::Full => Some(0),
        };
        Ok(())
    }

    pub(crate) fn mark_source_bound(&mut self) {
        self.container.mark_source_bound();
    }
}

impl<'a> Shape for Placeholder<'a>
where
    'a: 'static,
{
    fn properties(&self) -> &ShapeProperties {
        &self.container.properties
    }

    fn properties_mut(&mut self) -> Result<&mut ShapeProperties, super::shape::MutationError> {
        self.container.properties_mut_checked()
    }

    fn text(&self) -> super::super::package::Result<String> {
        Ok(self.container.text_content.clone().unwrap_or_default())
    }

    fn has_text(&self) -> bool {
        self.container
            .text_content
            .as_ref()
            .is_some_and(|text| !text.is_empty())
    }

    fn clone_box(&self) -> Box<dyn Shape> {
        Box::new(self.clone())
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::super::shape::ShapeType;
    use super::*;

    #[test]
    #[allow(
        clippy::field_reassign_with_default,
        reason = "the test builds a `ShapeProperties` fixture by mutating the default value"
    )]
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
    #[allow(
        clippy::field_reassign_with_default,
        reason = "the test builds a `ShapeProperties` fixture by mutating the default value"
    )]
    fn test_placeholder_type_operations() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::Placeholder;

        let mut placeholder = Placeholder::new(props, vec![]);
        assert!(
            placeholder
                .set_placeholder_type(PlaceholderType::Title)
                .is_ok()
        );
        assert!(placeholder.set_size(1).is_ok());
        assert!(placeholder.set_index(0).is_ok());

        assert_eq!(placeholder.placeholder_type(), PlaceholderType::Title);
        assert_eq!(placeholder.size(), Some(1));
        assert_eq!(placeholder.index(), Some(0));
        assert!(placeholder.is_title());
        assert!(!placeholder.is_content());
    }

    #[test]
    #[allow(
        clippy::field_reassign_with_default,
        reason = "the test builds a `ShapeProperties` fixture by mutating the default value"
    )]
    fn test_placeholder_media_check() {
        let mut props = ShapeProperties::default();
        props.shape_type = ShapeType::Placeholder;

        let mut placeholder = Placeholder::new(props, vec![]);

        // Test picture placeholder
        assert!(
            placeholder
                .set_placeholder_type(PlaceholderType::Picture)
                .is_ok()
        );
        assert!(placeholder.is_media());
        assert!(!placeholder.is_title());

        // Test title placeholder
        assert!(
            placeholder
                .set_placeholder_type(PlaceholderType::Title)
                .is_ok()
        );
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

    #[test]
    fn placeholder_size_uses_native_ppt_values() {
        let mut placeholder = Placeholder::new(ShapeProperties::default(), vec![]);

        for (raw, expected) in [
            (0, PlaceholderSize::Full),
            (1, PlaceholderSize::Half),
            (2, PlaceholderSize::Quarter),
        ] {
            assert!(placeholder.set_size(raw).is_ok());
            assert_eq!(placeholder.placeholder_size(), expected);
        }

        for (size, expected_raw) in [
            (PlaceholderSize::Full, 0),
            (PlaceholderSize::Half, 1),
            (PlaceholderSize::Quarter, 2),
        ] {
            assert!(placeholder.set_placeholder_size(size).is_ok());
            assert_eq!(placeholder.size(), Some(expected_raw));
        }
    }

    #[test]
    fn source_bound_placeholder_setters_are_atomic() {
        let mut placeholder = Placeholder::new(ShapeProperties::default(), Vec::new());
        let before_type = placeholder.placeholder_type();
        let before_size = placeholder.size();
        let before_index = placeholder.index();
        placeholder.container.mark_source_bound();
        let error = Err(super::super::shape::MutationError::SourceBound {
            mutation: super::super::shape::Mutation::Placeholder,
        });

        assert_eq!(
            placeholder.set_placeholder_type(PlaceholderType::Title),
            error
        );
        assert_eq!(placeholder.set_size(1), error);
        assert_eq!(placeholder.set_index(2), error);
        assert_eq!(
            placeholder.set_placeholder_size(PlaceholderSize::Quarter),
            error
        );
        assert_eq!(placeholder.placeholder_type(), before_type);
        assert_eq!(placeholder.size(), before_size);
        assert_eq!(placeholder.index(), before_index);
    }
}
