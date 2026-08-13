//! `OpenDocument` Presentation builder.
//!
//! This module provides a builder pattern for creating new ODP presentations from scratch.

mod edit;
mod package;
mod snapshot;
mod transition;
mod validation;
mod xml;

use crate::Reference;
use crate::Slide;
use crate::model::action::write_event_listeners;
use crate::model::animation::validate_animation_roots;
use crate::model::legacy_animation::validate_legacy_animation_root;
use crate::model::media::{EmbeddedMedia, embed_media};
use crate::model::slide::validate_z_index;
use litchi_core::{Metadata, Result, xml::escape_xml};
use litchi_odf_common::{
    drawing::authoring::Length,
    media::authoring::{allocate_picture_path, validate_payload},
};
use std::collections::{BTreeMap, BTreeSet};
use std::path::Path;

pub(crate) use self::transition::{
    DEFAULT_DRAWING_PAGE_STYLE, DEFAULT_DRAWING_PAGE_STYLE_NAME, generate_transition_styles,
    push_transition_style, slide_style_name,
};
use self::validation::{
    push_drawing_attributes, validate_drawing_shape_parent,
    validate_required_three_dimensional_attributes, validate_three_dimensional_child_order,
};
pub(crate) use self::xml::generate_text_paragraphs;
use self::xml::push_optional_attribute;

/// Builder for creating new ODP presentations.
///
/// This builder allows you to create ODP presentations programmatically by adding
/// slides with text and shapes, then saving them to a file or bytes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::Builder;
///
/// # fn main() -> litchi_core::Result<()> {
/// let mut builder = Builder::new();
/// builder.add_slide_with_title("Welcome", "This is my presentation")?;
/// builder.add_slide_with_title("Slide 2", "More content here")?;
/// builder.save("presentation.odp")?;
/// # Ok(())
/// # }
/// ```
pub struct Builder {
    slides: Vec<Slide>,
    metadata: Metadata,
    media_files: BTreeMap<String, EmbeddedMedia>,
    settings: Option<crate::Settings>,
    declarations: Option<crate::model::declaration::Collection>,
    page_metadata: Option<crate::model::page_metadata::Collection>,
    page_layouts: crate::model::page_layout::Collection,
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

impl Builder {
    fn generate_enhanced_geometry_xml(geometry: &crate::EnhancedGeometry) -> Result<String> {
        if geometry.children.len() > 65_536 {
            return Err(litchi_core::Error::InvalidFormat(
                "enhanced geometry exceeds 65536 equations and handles".to_string(),
            ));
        }
        let mut output = String::from("<draw:enhanced-geometry");
        push_drawing_attributes(&mut output, &geometry.attributes)?;
        if geometry.children.is_empty() {
            output.push_str("/>");
            return Ok(output);
        }
        output.push('>');
        let mut handle_seen = false;
        for child in &geometry.children {
            if child.kind == crate::EnhancedGeometryChildKind::Equation && handle_seen {
                return Err(litchi_core::Error::InvalidFormat(
                    "draw:equation cannot follow draw:handle".to_string(),
                ));
            }
            if child.kind == crate::EnhancedGeometryChildKind::Handle {
                handle_seen = true;
            }
            if child
                .attributes
                .iter()
                .any(|attribute| attribute.namespace != crate::DrawingAttributeNamespace::Drawing)
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "{} attributes must use the drawing namespace",
                    child.kind.element_name()
                )));
            }
            output.push('<');
            output.push_str(child.kind.element_name());
            push_drawing_attributes(&mut output, &child.attributes)?;
            output.push_str("/>");
        }
        output.push_str("</draw:enhanced-geometry>");
        Ok(output)
    }

    /// Create a new presentation builder
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::Builder;
    ///
    /// let builder = Builder::new();
    /// ```
    #[must_use]
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
            media_files: BTreeMap::new(),
            settings: None,
            declarations: None,
            page_metadata: None,
            page_layouts: crate::model::page_layout::Collection::default(),
        }
    }

    /// Return validated page-layout definitions that will be written to `styles.xml`.
    #[must_use]
    pub fn layouts(&self) -> &crate::model::page_layout::Collection {
        &self.page_layouts
    }

    /// Replace all custom page-layout definitions written by this builder.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn set_layouts(
        &mut self,
        layouts: crate::model::page_layout::Collection,
    ) -> Result<&mut Self> {
        layouts.validate()?;
        self.page_layouts = layouts;
        Ok(self)
    }

    /// Add one custom page layout without changing existing builder behavior.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn add_layout(&mut self, layout: crate::model::page_layout::Layout) -> Result<&mut Self> {
        let mut layouts = self.page_layouts.clone();
        layouts.layouts.push(layout);
        layouts.validate()?;
        self.page_layouts = layouts;
        Ok(self)
    }

    /// Return the inert slide-show settings.
    #[must_use]
    pub fn settings(&self) -> Option<&crate::Settings> {
        self.settings.as_ref()
    }

    /// Set or clear validated slide-show settings without executing them.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn set_settings(&mut self, settings: Option<crate::Settings>) -> Result<&mut Self> {
        if let Some(new_settings) = &settings {
            new_settings.validate()?;
        }
        self.settings = settings;
        Ok(self)
    }

    /// Return inert presentation declarations and page bindings.
    #[must_use]
    pub fn declarations(&self) -> Option<&crate::model::declaration::Collection> {
        self.declarations.as_ref()
    }

    /// Set or clear validated presentation declarations and page bindings.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn set_declarations(
        &mut self,
        declarations: Option<crate::model::declaration::Collection>,
    ) -> Result<&mut Self> {
        if let Some(new_declarations) = &declarations {
            new_declarations.validate()?;
        }
        self.declarations = declarations;
        Ok(self)
    }

    /// Return static page names, IDs, and layout/master references.
    #[must_use]
    pub fn pages(&self) -> Option<&crate::model::page_metadata::Collection> {
        self.page_metadata.as_ref()
    }

    /// Set or clear validated static page metadata.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn set_pages(
        &mut self,
        metadata: Option<crate::model::page_metadata::Collection>,
    ) -> Result<&mut Self> {
        if let Some(new_metadata) = &metadata {
            new_metadata.validate()?;
        }
        self.page_metadata = metadata;
        Ok(self)
    }

    /// Set document metadata
    ///
    /// # Arguments
    ///
    /// * `metadata` - Document metadata (title, author, etc.)
    pub fn set_metadata(&mut self, metadata: Metadata) {
        self.metadata = metadata;
    }

    /// Add a package-contained audio or video payload.
    ///
    /// The returned inert reference can be attached to a shape with
    /// [`crate::Shape::with_media`]. External resources are never fetched.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn embed_media(
        &mut self,
        path: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
        media_type: impl Into<String>,
    ) -> Result<Reference> {
        embed_media(&mut self.media_files, path, bytes, media_type)
    }

    /// Embed a PNG, JPEG, or GIF image and add it as a picture frame to a slide.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn insert_image(
        &mut self,
        slide_index: usize,
        bytes: &[u8],
        x: &Length,
        y: &Length,
        width: &Length,
        height: &Length,
    ) -> Result<&mut Self> {
        if slide_index >= self.slides.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "slide index {slide_index} is out of bounds"
            )));
        }

        let format = validate_payload(bytes)?;
        let extension = format.extension();
        let path = allocate_picture_path(extension, |candidate| {
            let stem = candidate
                .rsplit_once('.')
                .map_or(candidate, |(stem, _)| stem);
            self.media_files.keys().any(|existing| {
                existing.starts_with(stem) && existing.as_bytes().get(stem.len()) == Some(&b'.')
            })
        })?;

        // All fallible input validation is complete before staging the media.
        // The known-safe path, bounded bytes, and fixed MIME type leave no
        // fallible work between this insertion and adding the picture frame.
        embed_media(
            &mut self.media_files,
            path.clone(),
            bytes.to_vec(),
            format.media_type(),
        )?;

        let mut picture = crate::Shape::new().with_image_href(path);
        picture.x = Some(x.as_str().to_owned());
        picture.y = Some(y.as_str().to_owned());
        picture.width = Some(width.as_str().to_owned());
        picture.height = Some(height.as_str().to_owned());
        self.slides[slide_index].shapes.push(picture);
        Ok(self)
    }

    /// Add a slide with title and text content
    ///
    /// # Arguments
    ///
    /// * `title` - Title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_slide_with_title("Introduction", "Welcome to our presentation")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn add_slide_with_title(&mut self, title: &str, text: &str) -> Result<&mut Self> {
        Ok(self.append_slide(edit::titled_slide(self.slides.len(), title, text)))
    }

    /// Add a slide with only text content (no title)
    ///
    /// # Arguments
    ///
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_slide("Simple slide content")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn add_slide(&mut self, text: &str) -> Result<&mut Self> {
        Ok(self.append_slide(edit::text_slide(self.slides.len(), text)))
    }

    /// Add a Slide element directly
    ///
    /// # Arguments
    ///
    /// * `slide` - A complete `Slide` element to add
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::{Builder, Slide, Shape};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// let slide = Slide {
    ///     title: Some("Custom Slide".to_string()),
    ///     text: "Custom content".to_string(),
    ///     index: 0,
    ///     notes: Some("Speaker notes".to_string()),
    ///     transition: None,
    ///     animations: vec![],
    ///     legacy_animation: None,
    ///     shapes: vec![],
    /// };
    /// builder.add_slide_element(slide)?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn add_slide_element(&mut self, slide: Slide) -> Result<&mut Self> {
        Ok(self.append_slide(slide))
    }

    /// Generate XML for a shape
    pub(crate) fn generate_shape_xml(shape: &crate::Shape, idx: usize) -> Result<String> {
        let mut node_count = 0usize;
        Self::generate_shape_xml_at_depth(shape, idx, 0, None, &mut node_count)
    }

    #[allow(
        clippy::wildcard_enum_match_arm,
        reason = "ShapeType is a non-exhaustive litchi-core enum, so a wildcard arm is required and intentionally covers any future variants"
    )]
    fn generate_shape_xml_at_depth(
        shape: &crate::Shape,
        idx: usize,
        depth: usize,
        parent_kind: Option<crate::DrawingShapeKind>,
        node_count: &mut usize,
    ) -> Result<String> {
        use crate::DrawingShapeKind;
        use litchi_core::ShapeType;

        if depth > 64 {
            return Err(litchi_core::Error::InvalidFormat(
                "ODP shape groups exceed 64 levels".to_string(),
            ));
        }
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("ODP shape count overflow".to_string())
        })?;
        if *node_count > 65_536 {
            return Err(litchi_core::Error::InvalidFormat(
                "ODP document exceeds 65536 shapes".to_string(),
            ));
        }
        if shape.shape_type != ShapeType::Group && !shape.children.is_empty() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "non-group ODP shape '{}' cannot contain nested shapes",
                shape.name.as_deref().unwrap_or("unnamed")
            )));
        }

        let generated = shape.drawing_kind.is_none() && shape.shape_type != ShapeType::Group;
        let x = shape.x.as_deref().or(generated.then_some("2cm"));
        let y = shape.y.as_deref().or(generated.then_some("8cm"));
        let width = shape.width.as_deref().or(generated.then_some("10cm"));
        let height = shape.height.as_deref().or(generated.then_some("5cm"));
        let default_name = format!("Shape{}", idx + 1);
        let name = shape.name.as_deref().unwrap_or(&default_name);
        let mut position_attributes = String::new();
        push_optional_attribute(&mut position_attributes, "svg:x", x);
        push_optional_attribute(&mut position_attributes, "svg:y", y);
        push_optional_attribute(&mut position_attributes, "svg:width", width);
        push_optional_attribute(&mut position_attributes, "svg:height", height);
        let mut line_attributes = String::new();
        push_optional_attribute(
            &mut line_attributes,
            "svg:x1",
            shape.x.as_deref().or(generated.then_some("2cm")),
        );
        push_optional_attribute(
            &mut line_attributes,
            "svg:y1",
            shape.y.as_deref().or(generated.then_some("8cm")),
        );
        push_optional_attribute(
            &mut line_attributes,
            "svg:x2",
            shape.width.as_deref().or(generated.then_some("12cm")),
        );
        push_optional_attribute(
            &mut line_attributes,
            "svg:y2",
            shape.height.as_deref().or(generated.then_some("8cm")),
        );
        let mut shape_attributes = String::new();
        push_optional_attribute(
            &mut shape_attributes,
            "draw:name",
            shape.name.as_deref().or(generated.then_some(name)),
        );
        push_optional_attribute(
            &mut shape_attributes,
            "draw:style-name",
            shape.style_name.as_deref().or(generated.then_some("gr3")),
        );
        push_optional_attribute(
            &mut shape_attributes,
            "draw:layer",
            shape.layer.as_deref().or(generated.then_some("layout")),
        );
        if let Some(z_index) = &shape.z_index {
            validate_z_index(z_index)?;
            push_optional_attribute(&mut shape_attributes, "draw:z-index", Some(z_index));
        }
        if let Some(transform) = &shape.transform {
            push_optional_attribute(&mut shape_attributes, "draw:transform", Some(transform));
        }
        let presentation_class = shape
            .presentation_class
            .as_deref()
            .or_else(|| (shape.shape_type == ShapeType::Placeholder).then_some("object"));
        if let Some(class) = presentation_class {
            push_optional_attribute(&mut shape_attributes, "presentation:class", Some(class));
        }
        if let Some(placeholder) = shape.presentation_placeholder {
            push_optional_attribute(
                &mut shape_attributes,
                "presentation:placeholder",
                Some(if placeholder { "true" } else { "false" }),
            );
        }
        if let Some(user_transformed) = shape.presentation_user_transformed {
            push_optional_attribute(
                &mut shape_attributes,
                "presentation:user-transformed",
                Some(if user_transformed { "true" } else { "false" }),
            );
        }
        let mut drawing_attribute_names = BTreeSet::new();
        for attribute in &shape.drawing_attributes {
            let modeled = match attribute.namespace {
                crate::DrawingAttributeNamespace::Drawing => matches!(
                    attribute.local_name.as_str(),
                    "name" | "style-name" | "layer" | "z-index" | "transform"
                ),
                crate::DrawingAttributeNamespace::Svg => matches!(
                    attribute.local_name.as_str(),
                    "x" | "y" | "width" | "height" | "x1" | "y1" | "x2" | "y2"
                ),
                crate::DrawingAttributeNamespace::Dr3d
                | crate::DrawingAttributeNamespace::Table => false,
            };
            if modeled {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "ODP shape attribute '{}:{}' must use its dedicated shape field",
                    attribute.namespace.prefix(),
                    attribute.local_name
                )));
            }
            let qualified_name =
                format!("{}:{}", attribute.namespace.prefix(), attribute.local_name);
            if !drawing_attribute_names.insert(qualified_name.clone()) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "duplicate ODP shape attribute '{qualified_name}'"
                )));
            }
            shape_attributes.push(' ');
            shape_attributes.push_str(attribute.namespace.prefix());
            shape_attributes.push(':');
            shape_attributes.push_str(&attribute.local_name);
            shape_attributes.push_str("=\"");
            shape_attributes.push_str(&escape_xml(&attribute.value));
            shape_attributes.push('"');
        }

        // `draw:plugin` is only serialized inside `draw:frame`. Pictures qualify
        // because ODF allows a frame to pair a plugin with a fallback image;
        // every other shape type would silently drop the media reference.
        if shape.media.is_some()
            && !matches!(
                shape.shape_type,
                ShapeType::GraphicFrame | ShapeType::Picture
            )
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "ODP media shape '{name}' must use the graphic-frame or picture shape type"
            )));
        }

        let element_kind = match shape.shape_type {
            ShapeType::TextBox
            | ShapeType::Placeholder
            | ShapeType::Picture
            | ShapeType::GraphicFrame => DrawingShapeKind::Frame,
            ShapeType::AutoShape => shape.drawing_kind.unwrap_or(DrawingShapeKind::Rectangle),
            ShapeType::Line => shape.drawing_kind.unwrap_or(DrawingShapeKind::Line),
            ShapeType::Connector => shape.drawing_kind.unwrap_or(DrawingShapeKind::Connector),
            ShapeType::Group => shape.drawing_kind.unwrap_or(DrawingShapeKind::Group),
            _ => shape.drawing_kind.unwrap_or(DrawingShapeKind::Frame),
        };
        let compatible_kind = match shape.shape_type {
            ShapeType::TextBox
            | ShapeType::Placeholder
            | ShapeType::Picture
            | ShapeType::GraphicFrame => element_kind == DrawingShapeKind::Frame,
            ShapeType::AutoShape => !matches!(
                element_kind,
                DrawingShapeKind::Frame
                    | DrawingShapeKind::Line
                    | DrawingShapeKind::Measure
                    | DrawingShapeKind::Connector
                    | DrawingShapeKind::Group
                    | DrawingShapeKind::ThreeDimensionalScene
            ),
            ShapeType::Line => matches!(
                element_kind,
                DrawingShapeKind::Line | DrawingShapeKind::Measure
            ),
            ShapeType::Connector => element_kind == DrawingShapeKind::Connector,
            ShapeType::Group => matches!(
                element_kind,
                DrawingShapeKind::Group | DrawingShapeKind::ThreeDimensionalScene
            ),
            _ => true,
        };
        if !compatible_kind {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "ODP drawing element '{}' is incompatible with {} shape '{}'",
                element_kind.element_name(),
                shape.shape_type,
                name
            )));
        }
        validate_drawing_shape_parent(element_kind, parent_kind)?;
        validate_required_three_dimensional_attributes(element_kind, &shape.drawing_attributes)?;
        if element_kind.is_three_dimensional() && !shape.event_listeners.is_empty() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "3D shape '{name}' cannot contain presentation event listeners"
            )));
        }
        if parent_kind == Some(DrawingShapeKind::ThreeDimensionalScene) && shape.hyperlink.is_some()
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "3D scene child '{name}' cannot be wrapped in draw:a"
            )));
        }
        if shape.enhanced_geometry.is_some() && element_kind != DrawingShapeKind::CustomShape {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "enhanced geometry requires draw:custom-shape for shape '{name}'"
            )));
        }
        let element_name = element_kind.element_name();
        let mut xml = match shape.shape_type {
            ShapeType::TextBox | ShapeType::Placeholder => {
                if shape.has_text() {
                    format!(
                        r"<draw:frame{}{}><draw:text-box>{}</draw:text-box></draw:frame>",
                        shape_attributes,
                        position_attributes,
                        generate_text_paragraphs(&shape.text, Some("P2"))
                    )
                } else {
                    // Empty frame
                    format!(r"<draw:frame{shape_attributes}{position_attributes}/>")
                }
            },
            ShapeType::AutoShape => {
                if element_kind.is_three_dimensional() {
                    if shape.has_text()
                        || shape.image_href.is_some()
                        || shape.media.is_some()
                        || shape.enhanced_geometry.is_some()
                        || !shape.event_listeners.is_empty()
                    {
                        return Err(litchi_core::Error::InvalidFormat(format!(
                            "3D object '{name}' contains unsupported 2D shape payload"
                        )));
                    }
                    format!(r"<{element_name}{shape_attributes}{position_attributes}/>")
                } else {
                    let geometry = shape
                        .enhanced_geometry
                        .as_ref()
                        .map(Self::generate_enhanced_geometry_xml)
                        .transpose()?
                        .unwrap_or_default();
                    if shape.has_text() || !geometry.is_empty() {
                        let mut contents = if shape.has_text() {
                            generate_text_paragraphs(&shape.text, Some("P2"))
                        } else {
                            String::new()
                        };
                        contents.push_str(&geometry);
                        format!(
                            r"<{element_name}{shape_attributes}{position_attributes}>{contents}</{element_name}>"
                        )
                    } else {
                        format!(r"<{element_name}{shape_attributes}{position_attributes}/>")
                    }
                }
            },
            ShapeType::Picture => {
                // A plugin precedes its fallback image, matching how ODF
                // producers order the alternatives inside one frame.
                let mut contents = String::new();
                if let Some(media) = shape.media.as_ref() {
                    media.write_xml(&mut contents)?;
                }
                match shape.image_href.as_deref() {
                    Some(href) => {
                        contents.push_str(r#"<draw:image xlink:href=""#);
                        contents.push_str(&escape_xml(href));
                        contents.push_str(
                            r#"" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>"#,
                        );
                    },
                    None if contents.is_empty() => contents.push_str("<draw:image/>"),
                    None => {},
                }
                format!(
                    r"<draw:frame{shape_attributes}{position_attributes}>{contents}</draw:frame>"
                )
            },
            ShapeType::Line | ShapeType::Connector => {
                format!(r"<{element_name}{shape_attributes}{line_attributes}/>")
            },
            ShapeType::GraphicFrame if shape.media.is_some() => {
                let mut plugin = String::new();
                if let Some(media) = shape.media.as_ref() {
                    media.write_xml(&mut plugin)?;
                }
                format!(r"<draw:frame{shape_attributes}{position_attributes}>{plugin}</draw:frame>")
            },
            ShapeType::Group => {
                if shape.has_text()
                    || shape.image_href.is_some()
                    || shape.media.is_some()
                    || shape.enhanced_geometry.is_some()
                {
                    return Err(litchi_core::Error::InvalidFormat(format!(
                        "ODP group shape '{name}' contains non-group payload"
                    )));
                }
                let container_position = if element_kind == DrawingShapeKind::ThreeDimensionalScene
                {
                    position_attributes.as_str()
                } else {
                    ""
                };
                if shape.children.is_empty() {
                    format!(r"<{element_name}{shape_attributes}{container_position}/>")
                } else {
                    if element_kind == DrawingShapeKind::ThreeDimensionalScene {
                        validate_three_dimensional_child_order(&shape.children)?;
                    }
                    let mut children = String::new();
                    for (child_index, child) in shape.children.iter().enumerate() {
                        children.push_str(&Self::generate_shape_xml_at_depth(
                            child,
                            child_index,
                            depth + 1,
                            Some(element_kind),
                            node_count,
                        )?);
                    }
                    format!(
                        r"<{element_name}{shape_attributes}{container_position}>{children}</{element_name}>"
                    )
                }
            },
            ShapeType::Table | ShapeType::GraphicFrame | ShapeType::Unknown => {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "ODP serialization does not have enough data to write {} shape '{}'",
                    shape.shape_type, name
                )));
            },
            _ => {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "unsupported ODP shape type '{}' for shape '{}'",
                    shape.shape_type, name
                )));
            },
        };
        if !shape.event_listeners.is_empty() {
            let mut listeners = String::new();
            write_event_listeners(&mut listeners, &shape.event_listeners)?;
            let closing = format!("</{element_name}>");
            if xml.ends_with("/>") {
                xml.truncate(xml.len() - 2);
                xml.push('>');
                xml.push_str(&listeners);
                xml.push_str(&closing);
            } else if xml.ends_with(&closing) {
                let insertion = xml.len() - closing.len();
                xml.insert_str(insertion, &listeners);
            } else {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cannot attach ODP event listeners to shape '{name}'"
                )));
            }
        }
        if let Some(hyperlink) = &shape.hyperlink {
            let mut wrapped = String::with_capacity(xml.len() + 128);
            hyperlink.write_open_xml(&mut wrapped)?;
            wrapped.push_str(&xml);
            wrapped.push_str("</draw:a>");
            xml = wrapped;
        }
        Ok(xml)
    }

    pub(super) fn generate_notes_xml(notes: Option<&str>) -> String {
        notes
            .filter(|text| !text.is_empty())
            .map(|text| {
                format!(
                    r#"<presentation:notes><draw:frame draw:layer="layout" presentation:class="notes"><draw:text-box>{}</draw:text-box></draw:frame></presentation:notes>"#,
                    generate_text_paragraphs(text, None)
                )
            })
            .unwrap_or_default()
    }

    /// Generate the content.xml body for presentation
    fn generate_content_body(
        &self,
        extension_namespaces: &BTreeMap<String, String>,
    ) -> Result<String> {
        let shape_count = self.slides.iter().map(|s| s.shapes.len()).sum::<usize>();
        let mut estimated = 256usize;
        estimated += self.slides.len() * 128;
        estimated += shape_count * 192;
        estimated += self
            .slides
            .iter()
            .map(|s| s.text.len() + s.title.as_ref().map_or(0, String::len))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|sh| sh.text.len() + sh.name.as_ref().map_or(0, String::len))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        body.push_str(&crate::model::declaration::write_declaration_elements(
            self.declarations.as_ref(),
            self.slides.len(),
        )?);

        let page_names = crate::model::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        crate::model::settings::validate_page_references(self.settings.as_ref(), &page_names)?;

        for (i, slide) in self.slides.iter().enumerate() {
            let slide_style = slide_style_name(slide, i);
            if let Some(metadata) = &self.page_metadata {
                metadata.validate_for_slides(self.slides.len())?;
            }
            let page_attributes = crate::model::page_metadata::write_page_attributes(
                self.page_metadata.as_ref(),
                i,
                &slide_style,
            )?;
            let declaration_attributes = crate::model::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::model::declaration::Target::Slide,
            );
            body.push_str("<draw:page");
            body.push_str(&page_attributes);
            body.push_str(&declaration_attributes);
            body.push('>');

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                body.push_str(r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>"#);
                body.push_str(&generate_text_paragraphs(title, Some("P1")));
                body.push_str("</draw:text-box></draw:frame>");
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                body.push_str(r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y=""#);
                body.push_str(y_position);
                body.push_str(r#""><draw:text-box>"#);
                body.push_str(&generate_text_paragraphs(&slide.text, Some("P2")));
                body.push_str("</draw:text-box></draw:frame>");
            }

            // Add custom shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&Self::generate_shape_xml(shape, shape_idx)?);
            }

            for animation in &slide.animations {
                animation.write_xml(&mut body, extension_namespaces)?;
            }
            if let Some(animation) = &slide.legacy_animation {
                animation.write_xml(&mut body, extension_namespaces)?;
            }

            let notes_attributes = crate::model::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::model::declaration::Target::Notes,
            );
            body.push_str(&crate::model::declaration::apply_notes_binding(
                Self::generate_notes_xml(slide.notes.as_deref()),
                &notes_attributes,
            )?);

            body.push_str("</draw:page>");
        }

        body.push_str(&crate::model::settings::write(self.settings.as_ref())?);

        Ok(body)
    }

    /// Generate the complete content.xml for presentation
    fn generate_content_xml(&self) -> Result<String> {
        let mut extension_uris = BTreeSet::new();
        for slide in &self.slides {
            validate_animation_roots(&slide.animations)?;
            for animation in &slide.animations {
                animation.collect_extension_namespaces(&mut extension_uris);
            }
            if let Some(animation) = &slide.legacy_animation {
                validate_legacy_animation_root(animation)?;
                animation.collect_extension_namespaces(&mut extension_uris);
            }
        }
        let extension_namespaces = extension_uris
            .into_iter()
            .enumerate()
            .map(|(index, uri)| (uri, format!("anim-ext{}", index + 1)))
            .collect::<BTreeMap<_, _>>();
        let mut extension_declarations = String::new();
        for (uri, prefix) in &extension_namespaces {
            if uri.is_empty() {
                return Err(litchi_core::Error::InvalidFormat(
                    "animation extension namespace URI cannot be empty".to_string(),
                ));
            }
            extension_declarations.push_str(" xmlns:");
            extension_declarations.push_str(prefix);
            extension_declarations.push_str("=\"");
            extension_declarations.push_str(&escape_xml(uri));
            extension_declarations.push('"');
        }
        let body = self.generate_content_body(&extension_namespaces)?;
        let transition_styles = generate_transition_styles(&self.slides);

        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office"{extension_declarations} office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{transition_styles}</office:automatic-styles><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
        ))
    }

    /// Generate meta.xml with metadata
    fn generate_meta_xml(&self) -> String {
        let mut meta = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator>"#,
        );

        if let Some(created) = self.metadata.created.as_ref() {
            meta.push_str("<meta:creation-date>");
            meta.push_str(&created.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true));
            meta.push_str("</meta:creation-date>");
        } else if let Some(created) = self.metadata.created_local.as_ref() {
            meta.push_str("<meta:creation-date>");
            meta.push_str(&created.format("%Y-%m-%dT%H:%M:%S%.f").to_string());
            meta.push_str("</meta:creation-date>");
        }

        if let Some(modified) = self.metadata.modified.as_ref() {
            meta.push_str("<dc:date>");
            meta.push_str(&modified.to_rfc3339_opts(chrono::SecondsFormat::AutoSi, true));
            meta.push_str("</dc:date>");
        } else if let Some(modified) = self.metadata.modified_local.as_ref() {
            meta.push_str("<dc:date>");
            meta.push_str(&modified.format("%Y-%m-%dT%H:%M:%S%.f").to_string());
            meta.push_str("</dc:date>");
        }

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            meta.push_str("<dc:title>");
            meta.push_str(&escape_xml(title));
            meta.push_str("</dc:title>");
        }

        if let Some(ref author) = self.metadata.author {
            meta.push_str("<dc:creator>");
            meta.push_str(&escape_xml(author));
            meta.push_str("</dc:creator>");
        }

        meta.push_str("</office:meta>");
        meta.push_str("</office:document-meta>");

        meta
    }

    /// Build the presentation and return as bytes
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_slide("Slide content")?;
    /// let bytes = builder.build()?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn build(self) -> Result<Vec<u8>> {
        package::build(&self)
    }

    /// Build and save the presentation to a file
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::Builder;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut builder = Builder::new();
    /// builder.add_slide("Slide content")?;
    /// builder.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn save<P: AsRef<Path>>(self, path: P) -> Result<()> {
        let bytes = self.build()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }
}

#[cfg(test)]
#[allow(
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use crate::core::OwnedPackage;
    use crate::{
        Action, Direction, DrawingHyperlink, Effect, EventListener, HyperlinkShow, Parameter,
        Presentation, ScriptEventListener, Shape, ShapeEventListener, Sound, SoundShow, Speed,
        Style, Type,
    };
    use chrono::{NaiveDate, TimeZone, Utc};
    use litchi_core::ShapeType;

    #[test]
    fn fresh_builder_output_is_deterministic_without_ambient_timestamps() {
        let first = Builder::new().build().unwrap();
        let second = Builder::new().build().unwrap();
        assert_eq!(first, second);

        let package = OwnedPackage::from_bytes(first).unwrap();
        let meta = String::from_utf8(package.get_file("meta.xml").unwrap()).unwrap();
        assert!(!meta.contains("meta:creation-date"));
        assert!(!meta.contains("<dc:date>"));
    }

    #[test]
    fn builder_serializes_explicit_utc_and_local_metadata_dates() {
        let mut builder = Builder::new();
        builder.set_metadata(Metadata {
            title: Some("Title & details".to_string()),
            author: Some("Ada <Lovelace>".to_string()),
            created: Some(Utc.with_ymd_and_hms(2024, 1, 2, 3, 4, 5).single().unwrap()),
            modified_local: Some(
                NaiveDate::from_ymd_opt(2024, 6, 7)
                    .unwrap()
                    .and_hms_opt(8, 9, 10)
                    .unwrap(),
            ),
            ..Default::default()
        });

        let package = OwnedPackage::from_bytes(builder.build().unwrap()).unwrap();
        let meta = String::from_utf8(package.get_file("meta.xml").unwrap()).unwrap();
        assert!(meta.contains("<meta:creation-date>2024-01-02T03:04:05Z</meta:creation-date>"));
        assert!(meta.contains("<dc:date>2024-06-07T08:09:10</dc:date>"));
        assert!(meta.contains("<dc:title>Title &amp; details</dc:title>"));
        assert!(meta.contains("<dc:creator>Ada &lt;Lovelace&gt;</dc:creator>"));
    }

    #[test]
    fn writes_native_rectangles_and_connectors() {
        let mut rectangle = Shape::new();
        rectangle.name = Some("Box & label".to_string());
        rectangle.text = "Visible <text>".to_string();
        let rectangle_xml = Builder::generate_shape_xml(&rectangle, 0).unwrap();
        assert!(rectangle_xml.starts_with("<draw:rect"));
        assert!(rectangle_xml.contains("Visible &lt;text&gt;"));

        let mut connector = Shape::new();
        connector.shape_type = ShapeType::Connector;
        let connector_xml = Builder::generate_shape_xml(&connector, 1).unwrap();
        assert!(connector_xml.starts_with("<draw:connector"));
    }

    #[test]
    fn writes_explicit_drawing_kinds_and_escaped_geometry_attributes() {
        let mut shape = Shape::new();
        shape.name = Some("Ellipse".to_string());
        shape.drawing_kind = Some(crate::DrawingShapeKind::Ellipse);
        shape.drawing_attributes.push(
            crate::DrawingAttribute::new(
                crate::DrawingAttributeNamespace::Drawing,
                "kind",
                "section & arc",
            )
            .unwrap(),
        );
        shape.drawing_attributes.push(
            crate::DrawingAttribute::new(crate::DrawingAttributeNamespace::Svg, "rx", "2cm")
                .unwrap(),
        );
        let xml = Builder::generate_shape_xml(&shape, 0).unwrap();
        assert!(xml.starts_with("<draw:ellipse"));
        assert!(xml.contains(r#"draw:kind="section &amp; arc""#));
        assert!(xml.contains(r#"svg:rx="2cm""#));
        assert!(!xml.contains("svg:x="));

        shape.shape_type = ShapeType::Line;
        assert!(Builder::generate_shape_xml(&shape, 0).is_err());

        shape.shape_type = ShapeType::AutoShape;
        shape.drawing_attributes.push(
            crate::DrawingAttribute::new(crate::DrawingAttributeNamespace::Svg, "rx", "3cm")
                .unwrap(),
        );
        assert!(Builder::generate_shape_xml(&shape, 0).is_err());

        let mut reserved = Shape::new();
        reserved.drawing_attributes.push(
            crate::DrawingAttribute::new(
                crate::DrawingAttributeNamespace::Drawing,
                "name",
                "duplicate",
            )
            .unwrap(),
        );
        assert!(Builder::generate_shape_xml(&reserved, 0).is_err());
    }

    #[test]
    fn validates_three_dimensional_scene_hierarchy_and_light_order() {
        use crate::DrawingShapeKind;

        let mut light = Shape::new();
        light.drawing_kind = Some(DrawingShapeKind::ThreeDimensionalLight);
        light.drawing_attributes.push(
            crate::DrawingAttribute::new(
                crate::DrawingAttributeNamespace::Dr3d,
                "direction",
                "(0 0 -1)",
            )
            .unwrap(),
        );
        let mut cube = Shape::new();
        cube.drawing_kind = Some(DrawingShapeKind::ThreeDimensionalCube);

        let mut scene = Shape::new();
        scene.shape_type = ShapeType::Group;
        scene.drawing_kind = Some(DrawingShapeKind::ThreeDimensionalScene);
        scene.children = vec![light.clone(), cube.clone()];
        let xml = Builder::generate_shape_xml(&scene, 0).unwrap();
        assert!(xml.starts_with("<dr3d:scene"));
        assert!(xml.contains("<dr3d:light"));
        assert!(xml.contains("<dr3d:cube"));

        let mut missing_direction = Shape::new();
        missing_direction.drawing_kind = Some(DrawingShapeKind::ThreeDimensionalLight);
        scene.children = vec![missing_direction];
        assert!(Builder::generate_shape_xml(&scene, 0).is_err());

        scene.children = vec![cube.clone(), light];
        assert!(Builder::generate_shape_xml(&scene, 0).is_err());
        assert!(Builder::generate_shape_xml(&cube, 0).is_err());

        let mut group = Shape::new();
        group.shape_type = ShapeType::Group;
        group.drawing_kind = Some(DrawingShapeKind::Group);
        group.children.push(cube);
        assert!(Builder::generate_shape_xml(&group, 0).is_err());
    }

    #[test]
    fn writes_exact_shape_stacking_transform_and_presentation_role() {
        let mut shape = Shape::new();
        shape.layer = Some("foreground & controls".to_string());
        shape
            .set_z_index(Some("184467440737095516160".to_string()))
            .unwrap();
        shape.transform = Some("rotate (0.5)".to_string());
        shape.presentation_class = Some("chart & graph".to_string());
        shape.presentation_placeholder = Some(true);
        shape.presentation_user_transformed = Some(false);

        let xml = Builder::generate_shape_xml(&shape, 0).unwrap();
        assert!(xml.contains(r#"draw:layer="foreground &amp; controls""#));
        assert!(xml.contains(r#"draw:z-index="184467440737095516160""#));
        assert!(xml.contains(r#"draw:transform="rotate (0.5)""#));
        assert!(xml.contains(r#"presentation:class="chart &amp; graph""#));
        assert!(xml.contains(r#"presentation:placeholder="true""#));
        assert!(xml.contains(r#"presentation:user-transformed="false""#));
    }

    #[test]
    fn rejects_invalid_programmatic_shape_stacking_index() {
        let mut shape = Shape::new();
        assert!(shape.set_z_index(Some("-12".to_string())).is_err());

        shape.z_index = Some("not-an-integer".to_string());
        assert!(Builder::generate_shape_xml(&shape, 0).is_err());
    }

    #[test]
    fn rejects_shapes_without_enough_serializable_data() {
        for shape_type in [
            ShapeType::Table,
            ShapeType::GraphicFrame,
            ShapeType::Unknown,
        ] {
            let mut shape = Shape::new();
            shape.shape_type = shape_type;
            assert!(Builder::generate_shape_xml(&shape, 0).is_err());
        }
    }

    #[test]
    fn writes_nested_shape_groups_and_rejects_children_on_leaf_shapes() {
        let mut leaf = Shape::new();
        leaf.name = Some("Nested rectangle".to_string());
        leaf.text = "Nested text".to_string();

        let mut inner = Shape::new();
        inner.shape_type = ShapeType::Group;
        inner.drawing_kind = Some(crate::DrawingShapeKind::Group);
        inner.children.push(leaf.clone());

        let mut outer = Shape::new();
        outer.shape_type = ShapeType::Group;
        outer.drawing_kind = Some(crate::DrawingShapeKind::Group);
        outer.name = Some("Outer".to_string());
        outer.children.push(inner);

        let xml = Builder::generate_shape_xml(&outer, 0).unwrap();
        assert_eq!(xml.matches("<draw:g").count(), 2);
        assert!(xml.contains("<draw:rect"));
        assert!(xml.contains("Nested text"));

        leaf.children.push(Shape::new());
        assert!(Builder::generate_shape_xml(&leaf, 0).is_err());

        let mut too_deep = Shape::new();
        for _ in 0..66 {
            let mut parent = Shape::new();
            parent.shape_type = ShapeType::Group;
            parent.drawing_kind = Some(crate::DrawingShapeKind::Group);
            parent.children.push(too_deep);
            too_deep = parent;
        }
        assert!(Builder::generate_shape_xml(&too_deep, 0).is_err());
    }

    #[test]
    fn writes_paragraphs_and_explicit_odf_whitespace() {
        assert_eq!(
            generate_text_paragraphs(" leading  text\ttab\nnext ", Some("P&1")),
            r#"<text:p text:style-name="P&amp;1"><text:s/>leading<text:s text:c="2"/>text<text:tab/>tab</text:p><text:p text:style-name="P&amp;1">next<text:s/></text:p>"#
        );
    }

    #[test]
    fn transition_configuration_round_trips_through_a_package() {
        let mut builder = Builder::new();
        builder.add_slide("Transition slide").unwrap();
        let transition = builder.slides[0].transition_mut();
        transition
            .set_transition_type(Some(Type::Automatic))
            .set_style(Some(Style::new("fade-from-left").unwrap()))
            .set_speed(Some(Speed::Fast))
            .set_smil_type(Some("fade & dissolve"))
            .set_smil_subtype(Some("crossfade"))
            .set_direction(Some(Direction::Reverse));
        transition.set_fade_color(Some("#102030")).unwrap();
        transition.set_duration(Some("PT6.5S")).unwrap();
        let mut sound = Sound::new("Sounds/a&b.ogg");
        sound.play_full = Some(true);
        sound.actuate_on_request = true;
        sound.show = Some(SoundShow::Replace);
        sound.xml_id = Some("transitionSound1".to_string());
        transition.set_sound(Some(sound));

        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let slide = presentation.slides().unwrap().remove(0);
        let transition = slide.transition().unwrap();
        assert_eq!(transition.transition_type(), Some(Type::Automatic));
        assert_eq!(transition.style().unwrap().as_str(), "fade-from-left");
        assert_eq!(transition.speed(), Some(Speed::Fast));
        assert_eq!(transition.smil_type(), Some("fade & dissolve"));
        assert_eq!(transition.smil_subtype(), Some("crossfade"));
        assert_eq!(transition.direction(), Some(Direction::Reverse));
        assert_eq!(transition.fade_color(), Some("#102030"));
        assert_eq!(transition.duration(), Some("PT6.5S"));
        let sound = transition.sound().unwrap();
        assert_eq!(sound.href, "Sounds/a&b.ogg");
        assert_eq!(sound.play_full, Some(true));
        assert!(sound.actuate_on_request);
        assert_eq!(sound.show, Some(SoundShow::Replace));
        assert_eq!(sound.xml_id.as_deref(), Some("transitionSound1"));
    }

    #[test]
    fn embeds_and_round_trips_inert_presentation_media() {
        const VIDEO: &[u8] = b"test-video-payload";
        let mut builder = Builder::new();
        let mut media = builder
            .embed_media("Media/demo.mp4", VIDEO, "video/mp4")
            .unwrap();
        media
            .add_parameter(Parameter::new("autoplay", "false").unwrap())
            .unwrap();
        media.set_xml_id("demoVideo").unwrap();
        let shape = Shape::new().with_media(media.clone());
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: vec![shape],
            })
            .unwrap();

        let bytes = builder.build().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        assert_eq!(package.get_file("Media/demo.mp4").unwrap(), VIDEO);
        assert_eq!(
            package
                .package()
                .unwrap()
                .manifest()
                .get_media_type("Media/demo.mp4"),
            Some("video/mp4")
        );
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r#"draw:mime-type="video/mp4""#));
        assert!(content.contains(r#"xlink:href="Media/demo.mp4""#));
        assert!(content.contains(r#"draw:name="autoplay" draw:value="false""#));

        let presentation = Presentation::from_bytes(bytes).unwrap();
        let parsed = presentation.slides().unwrap();
        assert_eq!(parsed[0].shapes[0].media(), Some(&media));
        assert_eq!(
            presentation
                .media_data(parsed[0].shapes[0].media().unwrap())
                .unwrap(),
            Some(VIDEO.to_vec())
        );
    }

    #[test]
    fn shape_hyperlinks_and_actions_round_trip_through_a_package() {
        let mut hyperlink = DrawingHyperlink::new("#page2").unwrap();
        hyperlink.set_actuate_on_request(true);
        hyperlink.set_show(Some(HyperlinkShow::Replace));
        hyperlink
            .set_title(Some("Next & details".to_string()))
            .unwrap();
        hyperlink
            .set_xml_id(Some("actionLink1".to_string()))
            .unwrap();

        let mut action = EventListener::new("dom:click", Action::Sound).unwrap();
        action.effect = Some(Effect::new("fade").unwrap());
        action.speed = Some(Speed::Fast);
        let mut sound = Sound::new("Sounds/click.ogg");
        sound.play_full = Some(true);
        sound.actuate_on_request = true;
        sound.xml_id = Some("clickSound".to_string());
        action.sound = Some(sound);

        let mut shape = Shape::new();
        shape.name = Some("Action button".to_string());
        shape.text = "Continue".to_string();
        shape.set_hyperlink(Some(hyperlink.clone()));
        shape
            .add_event_listener(ShapeEventListener::Script(
                ScriptEventListener::macro_binding(
                    "dom:mouseover",
                    "ooo:script",
                    "Standard.Module1.Hover",
                )
                .unwrap(),
            ))
            .unwrap();
        shape
            .add_event_listener(ShapeEventListener::Action(Box::new(action)))
            .unwrap();

        let mut builder = Builder::new();
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: vec![shape],
            })
            .unwrap();
        let bytes = builder.build().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r##"<draw:a xlink:type="simple" xlink:href="#page2""##));
        assert!(content.contains(r#"presentation:action="sound""#));
        assert!(content.contains(r#"script:macro-name="Standard.Module1.Hover""#));

        let presentation = Presentation::from_bytes(bytes).unwrap();
        let slides = presentation.slides().unwrap();
        let parsed = &slides[0].shapes[0];
        assert_eq!(parsed.hyperlink(), Some(&hyperlink));
        assert_eq!(parsed.event_listeners().len(), 2);
        let ShapeEventListener::Action(action) = &parsed.event_listeners()[1] else {
            panic!("expected presentation action");
        };
        assert_eq!(action.action, Action::Sound);
        assert_eq!(action.sound.as_ref().unwrap().href, "Sounds/click.ogg");
    }

    #[test]
    fn rejects_unsafe_or_duplicate_embedded_media_paths() {
        let mut builder = Builder::new();
        assert!(
            builder
                .embed_media("../escape.mp4", [], "video/mp4")
                .is_err()
        );
        assert!(
            builder
                .embed_media("content.xml", [], "application/xml")
                .is_err()
        );
        builder
            .embed_media("Media/audio.ogg", [], "audio/ogg")
            .unwrap();
        assert!(
            builder
                .embed_media("Media/audio.ogg", [], "audio/ogg")
                .is_err()
        );
        assert!(
            builder
                .embed_media("Media/bad.bin", [], "not a mime type")
                .is_err()
        );
    }
}
