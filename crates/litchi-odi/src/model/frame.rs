//! Image-frame semantics.

use super::source::Source;

/// A semantic image frame.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Frame {
    name: Option<String>,
    source: Source,
    xml_id: Option<String>,
    title: Option<String>,
    description: Option<String>,
    media_type: Option<String>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
}

impl Frame {
    /// Creates a frame for the given image source without a name.
    #[must_use]
    pub fn new(source: Source) -> Self {
        Self {
            name: None,
            source,
            xml_id: None,
            title: None,
            description: None,
            media_type: None,
            x: None,
            y: None,
            width: None,
            height: None,
        }
    }

    /// Returns the frame name, if set.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the image payload source.
    #[must_use]
    pub fn source(&self) -> &Source {
        &self.source
    }

    /// Returns the stable XML identifier, if present.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Returns the short accessible title, if present.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    /// Returns the accessible description, if present.
    #[must_use]
    pub fn description(&self) -> Option<&str> {
        self.description.as_deref()
    }

    /// Returns the image's declared media type, if present.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Returns the lexical horizontal position, if present.
    #[must_use]
    pub fn x(&self) -> Option<&str> {
        self.x.as_deref()
    }

    /// Returns the lexical vertical position, if present.
    #[must_use]
    pub fn y(&self) -> Option<&str> {
        self.y.as_deref()
    }

    /// Returns the lexical frame width, if present.
    #[must_use]
    pub fn width(&self) -> Option<&str> {
        self.width.as_deref()
    }

    /// Returns the lexical frame height, if present.
    #[must_use]
    pub fn height(&self) -> Option<&str> {
        self.height.as_deref()
    }

    /// Sets the frame name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Sets the short accessible title.
    #[must_use]
    pub fn with_title(mut self, title: impl Into<String>) -> Self {
        self.title = Some(title.into());
        self
    }

    /// Sets the accessible description.
    #[must_use]
    pub fn with_description(mut self, description: impl Into<String>) -> Self {
        self.description = Some(description.into());
        self
    }

    pub(crate) fn from_scanned(source: Source, image: &litchi_odf_common::media::Image) -> Self {
        let mut frame = Self::new(source);
        frame.media_type.clone_from(&image.declared_media_type);
        if let Some(context) = &image.frame {
            frame.name.clone_from(&context.name);
            frame.xml_id.clone_from(&context.xml_id);
            frame.title.clone_from(&context.title);
            frame.description.clone_from(&context.description);
            frame.x.clone_from(&context.x);
            frame.y.clone_from(&context.y);
            frame.width.clone_from(&context.width);
            frame.height.clone_from(&context.height);
        }
        frame
    }
}
