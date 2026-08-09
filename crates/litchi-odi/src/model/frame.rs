//! Image-frame semantics.

use super::{map::ImageMap, source::Source};

/// A semantic image frame.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Frame {
    name: Option<String>,
    source: Source,
    xml_id: Option<String>,
    title: Option<String>,
    description: Option<String>,
    media_type: Option<String>,
    image_xml_id: Option<String>,
    filter_name: Option<String>,
    link_type: Option<String>,
    show: Option<String>,
    actuate: Option<String>,
    style_name: Option<String>,
    text_style_name: Option<String>,
    layer: Option<String>,
    z_index: Option<u32>,
    transform: Option<String>,
    anchor_type: Option<String>,
    relative_width: Option<String>,
    relative_height: Option<String>,
    copy_of: Option<String>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
    image_map: Option<ImageMap>,
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
            image_xml_id: None,
            filter_name: None,
            link_type: None,
            show: None,
            actuate: None,
            style_name: None,
            text_style_name: None,
            layer: None,
            z_index: None,
            transform: None,
            anchor_type: None,
            relative_width: None,
            relative_height: None,
            copy_of: None,
            x: None,
            y: None,
            width: None,
            height: None,
            image_map: None,
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

    /// Returns the `draw:image` XML identifier, if present.
    #[must_use]
    pub fn image_xml_id(&self) -> Option<&str> {
        self.image_xml_id.as_deref()
    }

    /// Returns the inert producer filter hint on `draw:image`.
    #[must_use]
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }

    /// Returns the image link type, if declared.
    #[must_use]
    pub fn link_type(&self) -> Option<&str> {
        self.link_type.as_deref()
    }

    /// Returns the image link presentation behavior, if declared.
    #[must_use]
    pub fn show(&self) -> Option<&str> {
        self.show.as_deref()
    }

    /// Returns the image link activation behavior, if declared.
    #[must_use]
    pub fn actuate(&self) -> Option<&str> {
        self.actuate.as_deref()
    }

    /// Returns the graphic style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Returns the paragraph style reference used by frame text.
    #[must_use]
    pub fn text_style_name(&self) -> Option<&str> {
        self.text_style_name.as_deref()
    }

    /// Returns the drawing layer name.
    #[must_use]
    pub fn layer(&self) -> Option<&str> {
        self.layer.as_deref()
    }

    /// Returns the non-negative stacking order.
    #[must_use]
    pub const fn z_index(&self) -> Option<u32> {
        self.z_index
    }

    /// Returns the lexical drawing transform.
    #[must_use]
    pub fn transform(&self) -> Option<&str> {
        self.transform.as_deref()
    }

    /// Returns the text anchoring mode.
    #[must_use]
    pub fn anchor_type(&self) -> Option<&str> {
        self.anchor_type.as_deref()
    }

    /// Returns the lexical relative width.
    #[must_use]
    pub fn relative_width(&self) -> Option<&str> {
        self.relative_width.as_deref()
    }

    /// Returns the lexical relative height.
    #[must_use]
    pub fn relative_height(&self) -> Option<&str> {
        self.relative_height.as_deref()
    }

    /// Returns the referenced source frame name for a copied frame.
    #[must_use]
    pub fn copy_of(&self) -> Option<&str> {
        self.copy_of.as_deref()
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

    /// Returns the optional client-side image map.
    #[must_use]
    pub const fn image_map(&self) -> Option<&ImageMap> {
        self.image_map.as_ref()
    }

    /// Sets the frame name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Sets the stable XML identifier on `draw:frame`.
    #[must_use]
    pub fn with_xml_id(mut self, value: impl Into<String>) -> Self {
        self.xml_id = Some(value.into());
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

    /// Sets the declared media type on `draw:image`.
    #[must_use]
    pub fn with_media_type(mut self, value: impl Into<String>) -> Self {
        self.media_type = Some(value.into());
        self
    }

    /// Sets the stable XML identifier on `draw:image`.
    #[must_use]
    pub fn with_image_xml_id(mut self, value: impl Into<String>) -> Self {
        self.image_xml_id = Some(value.into());
        self
    }

    /// Sets the inert producer filter hint on `draw:image`.
    #[must_use]
    pub fn with_filter_name(mut self, value: impl Into<String>) -> Self {
        self.filter_name = Some(value.into());
        self
    }

    /// Sets the lexical `XLink` type on `draw:image`.
    #[must_use]
    pub fn with_link_type(mut self, value: impl Into<String>) -> Self {
        self.link_type = Some(value.into());
        self
    }

    /// Sets the lexical `XLink` presentation behavior on `draw:image`.
    #[must_use]
    pub fn with_show(mut self, value: impl Into<String>) -> Self {
        self.show = Some(value.into());
        self
    }

    /// Sets the lexical `XLink` activation behavior on `draw:image`.
    #[must_use]
    pub fn with_actuate(mut self, value: impl Into<String>) -> Self {
        self.actuate = Some(value.into());
        self
    }

    /// Sets the graphic style reference.
    #[must_use]
    pub fn with_style_name(mut self, value: impl Into<String>) -> Self {
        self.style_name = Some(value.into());
        self
    }

    /// Sets the paragraph style reference used by frame text.
    #[must_use]
    pub fn with_text_style_name(mut self, value: impl Into<String>) -> Self {
        self.text_style_name = Some(value.into());
        self
    }

    /// Sets the drawing layer.
    #[must_use]
    pub fn with_layer(mut self, value: impl Into<String>) -> Self {
        self.layer = Some(value.into());
        self
    }

    /// Sets the non-negative stacking order.
    #[must_use]
    pub const fn with_z_index(mut self, value: u32) -> Self {
        self.z_index = Some(value);
        self
    }

    /// Sets the lexical drawing transform.
    #[must_use]
    pub fn with_transform(mut self, value: impl Into<String>) -> Self {
        self.transform = Some(value.into());
        self
    }

    /// Sets the text anchor mode.
    #[must_use]
    pub fn with_anchor_type(mut self, value: impl Into<String>) -> Self {
        self.anchor_type = Some(value.into());
        self
    }

    /// Sets lexical position and size values.
    #[must_use]
    pub fn with_geometry(
        mut self,
        x: impl Into<String>,
        y: impl Into<String>,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Self {
        self.x = Some(x.into());
        self.y = Some(y.into());
        self.width = Some(width.into());
        self.height = Some(height.into());
        self
    }

    /// Sets lexical relative width and height values.
    #[must_use]
    pub fn with_relative_size(
        mut self,
        width: impl Into<String>,
        height: impl Into<String>,
    ) -> Self {
        self.relative_width = Some(width.into());
        self.relative_height = Some(height.into());
        self
    }

    /// Sets the referenced source frame for `draw:copy-of`.
    #[must_use]
    pub fn with_copy_of(mut self, value: impl Into<String>) -> Self {
        self.copy_of = Some(value.into());
        self
    }

    /// Attaches a client-side image map.
    #[must_use]
    pub fn with_image_map(mut self, value: ImageMap) -> Self {
        self.image_map = Some(value);
        self
    }

    pub(crate) fn from_scanned(source: Source, image: &litchi_odf_common::media::Image) -> Self {
        let mut frame = Self::new(source);
        frame.media_type.clone_from(&image.declared_media_type);
        frame.image_xml_id.clone_from(&image.xml_id);
        frame.filter_name.clone_from(&image.filter_name);
        frame.link_type.clone_from(&image.link_type);
        frame.show.clone_from(&image.show);
        frame.actuate.clone_from(&image.actuate);
        if let Some(context) = &image.frame {
            frame.name.clone_from(&context.name);
            frame.xml_id.clone_from(&context.xml_id);
            frame.title.clone_from(&context.title);
            frame.description.clone_from(&context.description);
            frame.anchor_type.clone_from(&context.anchor_type);
            frame.x.clone_from(&context.x);
            frame.y.clone_from(&context.y);
            frame.width.clone_from(&context.width);
            frame.height.clone_from(&context.height);
        }
        frame
    }

    pub(crate) fn apply_properties(&mut self, properties: Properties) {
        self.style_name = properties.style_name;
        self.text_style_name = properties.text_style_name;
        self.layer = properties.layer;
        self.z_index = properties.z_index;
        self.transform = properties.transform;
        self.relative_width = properties.relative_width;
        self.relative_height = properties.relative_height;
        self.copy_of = properties.copy_of;
        self.image_map = properties.image_map;
    }

    pub(crate) fn set_name(&mut self, value: Option<String>) {
        self.name = value;
    }

    pub(crate) fn set_source(&mut self, value: Source) {
        self.source = value;
    }

    pub(crate) fn set_xml_id(&mut self, value: Option<String>) {
        self.xml_id = value;
    }

    pub(crate) fn set_title(&mut self, value: Option<String>) {
        self.title = value;
    }

    pub(crate) fn set_description(&mut self, value: Option<String>) {
        self.description = value;
    }

    pub(crate) fn set_media_type(&mut self, value: Option<String>) {
        self.media_type = value;
    }

    pub(crate) fn set_image_xml_id(&mut self, value: Option<String>) {
        self.image_xml_id = value;
    }

    pub(crate) fn set_filter_name(&mut self, value: Option<String>) {
        self.filter_name = value;
    }

    pub(crate) fn set_link_type(&mut self, value: Option<String>) {
        self.link_type = value;
    }

    pub(crate) fn set_show(&mut self, value: Option<String>) {
        self.show = value;
    }

    pub(crate) fn set_actuate(&mut self, value: Option<String>) {
        self.actuate = value;
    }

    pub(crate) fn set_style_name(&mut self, value: Option<String>) {
        self.style_name = value;
    }

    pub(crate) fn set_text_style_name(&mut self, value: Option<String>) {
        self.text_style_name = value;
    }

    pub(crate) fn set_layer(&mut self, value: Option<String>) {
        self.layer = value;
    }

    pub(crate) fn set_z_index(&mut self, value: Option<u32>) {
        self.z_index = value;
    }

    pub(crate) fn set_transform(&mut self, value: Option<String>) {
        self.transform = value;
    }

    pub(crate) fn set_anchor_type(&mut self, value: Option<String>) {
        self.anchor_type = value;
    }

    pub(crate) fn set_geometry(
        &mut self,
        x: Option<String>,
        y: Option<String>,
        width: Option<String>,
        height: Option<String>,
    ) {
        self.x = x;
        self.y = y;
        self.width = width;
        self.height = height;
    }

    pub(crate) fn set_relative_size(&mut self, width: Option<String>, height: Option<String>) {
        self.relative_width = width;
        self.relative_height = height;
    }

    pub(crate) fn set_image_map(&mut self, value: Option<ImageMap>) {
        self.image_map = value;
    }

    pub(crate) fn set_x(&mut self, value: Option<String>) {
        self.x = value;
    }

    pub(crate) fn set_y(&mut self, value: Option<String>) {
        self.y = value;
    }

    pub(crate) fn set_width(&mut self, value: Option<String>) {
        self.width = value;
    }

    pub(crate) fn set_height(&mut self, value: Option<String>) {
        self.height = value;
    }

    pub(crate) fn set_relative_width(&mut self, value: Option<String>) {
        self.relative_width = value;
    }

    pub(crate) fn set_relative_height(&mut self, value: Option<String>) {
        self.relative_height = value;
    }

    pub(crate) fn set_copy_of(&mut self, value: Option<String>) {
        self.copy_of = value;
    }
}

/// Additional ODF 1.4 frame semantics scanned by the ODI owner.
#[derive(Clone, Debug, Default)]
pub(crate) struct Properties {
    pub(crate) style_name: Option<String>,
    pub(crate) text_style_name: Option<String>,
    pub(crate) layer: Option<String>,
    pub(crate) z_index: Option<u32>,
    pub(crate) transform: Option<String>,
    pub(crate) relative_width: Option<String>,
    pub(crate) relative_height: Option<String>,
    pub(crate) copy_of: Option<String>,
    pub(crate) image_map: Option<ImageMap>,
}
