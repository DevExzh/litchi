//! Mutable presentation structure for in-place modifications.
//!
//! This module provides a mutable wrapper around ODP presentations that allows
//! for in-place modification of slides, shapes, and content.

use crate::core::{MetaXmlPatch, Structure, OwnedPackage, PackageWriter, patch_meta_xml};
use crate::animation::validate_animation_roots;
use crate::codec::content_source::ContentSource;
use crate::legacy_animation::validate_legacy_animation_root;
use crate::media::{EmbeddedMedia, embed_media, validate_package_media_path};
use crate::{MediaReference, Presentation, Shape, Slide};
use litchi_core::{Metadata, Result, xml::escape_xml};
use std::collections::{BTreeMap, BTreeSet};
use std::path::Path;

/// A mutable ODP presentation that supports in-place modifications.
///
/// This struct wraps an ODP presentation and provides methods to modify its content,
/// including adding, updating, and removing slides and shapes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::{Presentation, MutablePresentation};
///
/// # fn main() -> litchi_core::Result<()> {
/// // Open an existing presentation
/// let presentation = Presentation::open("input.odp")?;
/// let mut mutable = MutablePresentation::from_presentation(presentation)?;
///
/// // Modify the presentation
/// mutable.add_slide("New Slide", "Slide content")?;
/// mutable.remove_slide(0)?;
///
/// // Save the modified presentation
/// mutable.save("output.odp")?;
/// # Ok(())
/// # }
/// ```
pub struct MutablePresentation {
    /// Mutable slides
    slides: Vec<Slide>,
    /// Document metadata
    metadata: Metadata,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is)
    styles_xml: Option<String>,
    /// Original package retained for copying auxiliary package parts.
    source_package: Option<OwnedPackage>,
    /// Newly embedded package media, keyed by package path.
    media_files: BTreeMap<String, EmbeddedMedia>,
    /// Monotonic counter for authored frame names (1-based).
    next_frame_number: usize,
    /// Inert slide-show settings and custom shows.
    settings: Option<crate::Settings>,
    /// Inert header/footer/date-time declarations and page bindings.
    declarations: Option<crate::Declarations>,
    /// Static page names, IDs, and layout/master references.
    page_metadata: Option<crate::PageMetadataCollection>,
    /// Verbatim fragments of the source `content.xml`, when one was opened.
    ///
    /// Slides the caller never touched are re-emitted from these fragments so
    /// constructs outside the model — nested tables, image text alternatives,
    /// automatic styles, font declarations — survive a save.
    content_source: Option<ContentSource>,
    /// Slides exactly as parsed, used to detect which pages are still pristine.
    source_slides: Vec<Slide>,
    /// Declarations exactly as parsed; changing them rewrites every page.
    source_declarations: Option<crate::Declarations>,
    /// Page metadata exactly as parsed; changing it rewrites every page.
    source_page_metadata: Option<crate::PageMetadataCollection>,
}

/// Largest source slide count that is linearly rescanned for a match.
///
/// Beyond this, only the slide's own position is checked, keeping the retained
/// page lookup linear in the slide count.
const MAX_SOURCE_PAGE_SCAN: usize = 4_096;

/// Compare two slides ignoring their position in the deck.
///
/// Removing or reordering slides renumbers [`Slide::index`] without changing
/// the authored markup, so position is excluded from the comparison. The
/// destructuring binding makes any future field a compile error rather than a
/// silently ignored difference.
fn slide_content_eq(left: &Slide, right: &Slide) -> bool {
    let Slide {
        title,
        text,
        index: _,
        notes,
        transition,
        animations,
        legacy_animation,
        shapes,
    } = left;
    *title == right.title
        && *text == right.text
        && *notes == right.notes
        && *transition == right.transition
        && *animations == right.animations
        && *legacy_animation == right.legacy_animation
        && *shapes == right.shapes
}

impl MutablePresentation {
    /// Create a mutable presentation from an existing Presentation.
    ///
    /// This parses the presentation structure into mutable elements.
    ///
    /// # Arguments
    ///
    /// * `presentation` - The presentation to make mutable
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::{Presentation, MutablePresentation};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let mut mutable = MutablePresentation::from_presentation(presentation)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn from_presentation(presentation: Presentation) -> Result<Self> {
        let slides = presentation.slides()?;
        let metadata = presentation.metadata()?;
        let settings = presentation.settings()?;
        let declarations = presentation.declarations()?;
        let page_metadata = presentation.page_metadata()?;
        let mimetype = "application/vnd.oasis.opendocument.presentation".to_string();

        let styles_xml = presentation.styles_xml().map(str::to_owned);
        let content_source = ContentSource::parse(presentation.content_xml())?;
        let source_package = Some(presentation.into_package());

        let declarations = (!declarations.is_empty()).then_some(declarations);
        let page_metadata = (!page_metadata.is_empty()).then_some(page_metadata);
        let source_slides = match &content_source {
            // Only pages the scanner actually recovered can be retained; a
            // mismatch means the model and the markup disagree, so fall back to
            // regenerating every page.
            Some(source) if source.page_count() == slides.len() => slides.clone(),
            _ => Vec::new(),
        };
        let content_source = content_source.filter(|_| !source_slides.is_empty());

        Ok(Self {
            slides,
            metadata,
            mimetype,
            styles_xml,
            source_package,
            media_files: BTreeMap::new(),
            next_frame_number: 1,
            settings,
            source_declarations: declarations.clone(),
            declarations,
            source_page_metadata: page_metadata.clone(),
            page_metadata,
            content_source,
            source_slides,
        })
    }

    /// Create a new empty mutable presentation.
    ///
    /// # Examples
    ///
    /// ```
    /// use litchi_odp::MutablePresentation;
    ///
    /// let presentation = MutablePresentation::new();
    /// ```
    pub fn new() -> Self {
        Self {
            slides: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.presentation".to_string(),
            styles_xml: None,
            source_package: None,
            media_files: BTreeMap::new(),
            next_frame_number: 1,
            settings: None,
            declarations: None,
            page_metadata: None,
            content_source: None,
            source_slides: Vec::new(),
            source_declarations: None,
            source_page_metadata: None,
        }
    }

    /// Inspect named drawing fill-image definitions from current styles metadata.
    ///
    /// Links remain stored metadata: this does not follow them, load linked
    /// resources, or render images.
    pub fn drawing_fill_images(&self) -> Result<crate::drawing_fill_image::OdfDrawingFillImages> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_fill_image::parse_drawing_fill_images,
        )
    }

    /// Inspect named legacy and SVG drawing gradients from current styles metadata.
    ///
    /// This does not resolve style use sites or render gradients.
    pub fn drawing_gradients(&self) -> Result<crate::drawing_gradient::OdfDrawingGradients> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_gradient::parse_drawing_gradients,
        )
    }

    /// Inspect named drawing hatch definitions from current styles metadata.
    ///
    /// This does not resolve style use sites or render hatches.
    pub fn drawing_hatches(&self) -> Result<crate::drawing_hatch::OdfDrawingHatches> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_hatch::parse_drawing_hatches,
        )
    }

    /// Inspect named drawing marker definitions from current styles metadata.
    ///
    /// This does not resolve style use sites or render marker paths.
    pub fn drawing_markers(&self) -> Result<crate::drawing_marker::OdfDrawingMarkers> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_marker::parse_drawing_markers,
        )
    }

    /// Inspect named drawing opacity definitions from current styles metadata.
    ///
    /// This does not resolve style use sites or render opacity gradients.
    pub fn drawing_opacities(&self) -> Result<crate::drawing_opacity::OdfDrawingOpacities> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_opacity::parse_drawing_opacities,
        )
    }

    /// Inspect named drawing stroke-dash definitions from current styles metadata.
    ///
    /// This does not resolve style use sites or render strokes.
    pub fn drawing_stroke_dashes(
        &self,
    ) -> Result<crate::drawing_stroke_dash::OdfDrawingStrokeDashes> {
        self.styles_xml.as_deref().map_or_else(
            || Ok(Default::default()),
            crate::drawing_stroke_dash::parse_drawing_stroke_dashes,
        )
    }

    /// Return the inert slide-show settings.
    pub fn settings(&self) -> Option<&crate::Settings> {
        self.settings.as_ref()
    }

    /// Mutably access inert slide-show settings.
    pub fn settings_mut(&mut self) -> Option<&mut crate::Settings> {
        self.settings.as_mut()
    }

    /// Set or clear validated slide-show settings without executing them.
    pub fn set_settings(
        &mut self,
        settings: Option<crate::Settings>,
    ) -> Result<()> {
        if let Some(settings) = &settings {
            settings.validate()?;
        }
        self.settings = settings;
        Ok(())
    }

    /// Return inert presentation declarations and page bindings.
    pub fn declarations(&self) -> Option<&crate::Declarations> {
        self.declarations.as_ref()
    }

    /// Mutably access presentation declarations and page bindings.
    pub fn declarations_mut(&mut self) -> Option<&mut crate::Declarations> {
        self.declarations.as_mut()
    }

    /// Set or clear validated presentation declarations and page bindings.
    pub fn set_declarations(
        &mut self,
        declarations: Option<crate::Declarations>,
    ) -> Result<()> {
        if let Some(declarations) = &declarations {
            declarations.validate()?;
        }
        self.declarations = declarations;
        Ok(())
    }

    /// Return static page names, IDs, and layout/master references.
    pub fn page_metadata(&self) -> Option<&crate::PageMetadataCollection> {
        self.page_metadata.as_ref()
    }

    /// Mutably access static page metadata.
    pub fn page_metadata_mut(
        &mut self,
    ) -> Option<&mut crate::PageMetadataCollection> {
        self.page_metadata.as_mut()
    }

    /// Set or clear validated static page metadata.
    pub fn set_page_metadata(
        &mut self,
        metadata: Option<crate::PageMetadataCollection>,
    ) -> Result<()> {
        if let Some(metadata) = &metadata {
            metadata.validate()?;
        }
        self.page_metadata = metadata;
        Ok(())
    }

    /// Get all slides in the presentation.
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Get a mutable reference to all slides.
    pub fn slides_mut(&mut self) -> &mut Vec<Slide> {
        &mut self.slides
    }

    /// Get the presentation metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the presentation metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }

    /// Add a package-contained audio or video payload.
    ///
    /// Existing source-package paths cannot be replaced implicitly. The
    /// returned inert reference can be attached with [`Shape::with_media`].
    pub fn embed_media(
        &mut self,
        path: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
        media_type: impl Into<String>,
    ) -> Result<MediaReference> {
        let path = path.into();
        validate_package_media_path(&path)?;
        if let Some(package) = &self.source_package
            && package.has_file(&path)?
        {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "cannot replace existing ODP package media path '{path}' implicitly"
            )));
        }
        embed_media(&mut self.media_files, path, bytes, media_type)
    }

    /// Add a new slide to the end of the presentation.
    ///
    /// # Arguments
    ///
    /// * `title` - Optional title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content for slide 1")?;
    /// presentation.add_slide("Slide 2", "Content for slide 2")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_slide(&mut self, title: &str, text: &str) -> Result<()> {
        let slide = Slide {
            title: Some(title.to_string()),
            text: text.to_string(),
            index: self.slides.len(),
            notes: None,
            transition: None,
            animations: Vec::new(),
            legacy_animation: None,
            shapes: Vec::new(),
        };
        self.slides.push(slide);
        Ok(())
    }

    /// Insert a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Position to insert at (0-based)
    /// * `title` - Optional title for the slide
    /// * `text` - Text content for the slide
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("First", "Content 1")?;
    /// presentation.add_slide("Third", "Content 3")?;
    /// presentation.insert_slide(1, "Second", "Content 2")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index <= self.slides.len() {
            let candidate_metadata = super::page_metadata::metadata_after_page_insert(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = super::page_metadata::effective_page_names(
                Some(&candidate_metadata),
                self.slides.len() + 1,
            )?;
            super::settings::validate_page_references(
                self.settings.as_ref(),
                &candidate_names,
            )?;
            let slide = Slide {
                title: Some(title.to_string()),
                text: text.to_string(),
                index,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: Vec::new(),
            };
            self.slides.insert(index, slide);
            self.page_metadata = Some(candidate_metadata);

            // Update indices of subsequent slides
            for i in (index + 1)..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Remove a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to remove (0-based)
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content 1")?;
    /// presentation.add_slide("Slide 2", "Content 2")?;
    /// presentation.remove_slide(0)?; // Remove first slide
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_slide(&mut self, index: usize) -> Result<Slide> {
        if index < self.slides.len() {
            let current_names = super::page_metadata::effective_page_names(
                self.page_metadata.as_ref(),
                self.slides.len(),
            )?;
            let removed_name = &current_names[index];
            if super::settings::settings_reference_page(self.settings.as_ref(), removed_name) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cannot remove presentation page '{removed_name}' because slide-show settings reference it"
                )));
            }
            let candidate_metadata = super::page_metadata::metadata_after_page_remove(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = super::page_metadata::effective_page_names(
                candidate_metadata.as_ref(),
                self.slides.len() - 1,
            )?;
            super::settings::validate_page_references(
                self.settings.as_ref(),
                &candidate_names,
            )?;
            let slide = self.slides.remove(index);
            self.page_metadata = candidate_metadata;

            // Update indices of subsequent slides
            for i in index..self.slides.len() {
                self.slides[i].index = i;
            }

            Ok(slide)
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Update a slide at a specific index.
    ///
    /// # Arguments
    ///
    /// * `index` - Index of the slide to update (0-based)
    /// * `title` - New title for the slide
    /// * `text` - New text content
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Old Title", "Old content")?;
    /// presentation.update_slide(0, "New Title", "New content")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn update_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index < self.slides.len() {
            self.slides[index].title = Some(title.to_string());
            self.slides[index].text = text.to_string();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Index {} out of bounds (length: {})",
                index,
                self.slides.len()
            )))
        }
    }

    /// Clear all slides from the presentation.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content 1")?;
    /// presentation.add_slide("Slide 2", "Content 2")?;
    /// presentation.clear_slides();
    /// assert_eq!(presentation.slides().len(), 0);
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_slides(&mut self) {
        self.slides.clear();
    }

    /// Add a shape to a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide to add the shape to
    /// * `shape` - Shape to add
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::{MutablePresentation, Shape};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// let mut shape = Shape::new();
    /// shape.text = "Shape text".to_string();
    /// presentation.add_shape(0, shape)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn add_shape(&mut self, slide_index: usize, shape: Shape) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].shapes.push(shape);
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Insert a package-stored image onto a slide.
    ///
    /// The payload is sniffed (PNG, JPEG, and GIF are accepted), stored
    /// verbatim under `Pictures/` in the package with a manifest entry, and
    /// referenced from a `draw:frame`/`draw:image` element on the slide with
    /// the given `svg:x`/`svg:y`/`svg:width`/`svg:height` geometry. Picture
    /// numbering is global across the whole package, including pictures
    /// already present in a source document.
    ///
    /// Returns the allocated package path of the picture part.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::{MutablePresentation, OdfLength};
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Title", "Body")?;
    /// let png = b"\x89PNG\r\n\x1a\n".as_slice();
    /// let path = presentation.insert_image(
    ///     0,
    ///     png,
    ///     &OdfLength::centimeters(2.0),
    ///     &OdfLength::centimeters(3.0),
    ///     &OdfLength::centimeters(10.0),
    ///     &OdfLength::centimeters(5.0),
    /// )?;
    /// assert!(path.starts_with("Pictures/"));
    /// # Ok(())
    /// # }
    /// ```
    pub fn insert_image(
        &mut self,
        slide_index: usize,
        image: &[u8],
        x: &crate::odt::OdfLength,
        y: &crate::odt::OdfLength,
        width: &crate::odt::OdfLength,
        height: &crate::odt::OdfLength,
    ) -> Result<String> {
        use crate::odt::frame;
        if slide_index >= self.slides.len() {
            return Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )));
        }
        let format = frame::validate_image_payload(image)?;
        let path = frame::allocate_picture_path(format.extension(), |candidate| {
            // Picture numbering is global: a stem taken by any supported
            // extension blocks the whole index.
            let taken = |path: &str| {
                self.media_files.contains_key(path)
                    || self
                        .source_package
                        .as_ref()
                        .is_some_and(|package| package.has_file(path).unwrap_or(false))
                    || self
                        .slides
                        .iter()
                        .flat_map(|slide| slide.shapes.iter())
                        .any(|shape| shape.image_href() == Some(path))
            };
            if taken(candidate) {
                return true;
            }
            let stem = candidate.trim_end_matches(format.extension());
            ["png", "jpg", "gif"]
                .iter()
                .any(|extension| taken(&format!("{stem}{extension}")))
        })?;

        let name = format!("Image {}", self.next_frame_number);
        let shape = Shape {
            drawing_kind: Some(crate::DrawingShapeKind::Frame),
            name: Some(name),
            x: Some(x.as_str().to_string()),
            y: Some(y.as_str().to_string()),
            width: Some(width.as_str().to_string()),
            height: Some(height.as_str().to_string()),
            ..Shape::new()
        }
        .with_image_href(path.clone());
        self.add_shape(slide_index, shape)?;
        if let Err(error) = self.embed_media(path.clone(), image.to_vec(), format.media_type()) {
            // Roll the shape back so a failed insert leaves no trace.
            self.slides[slide_index].shapes.pop();
            return Err(error);
        }
        self.next_frame_number += 1;
        Ok(path)
    }

    /// Remove a shape from a slide.
    ///
    /// # Arguments
    ///
    /// * `slide_index` - Index of the slide
    /// * `shape_index` - Index of the shape to remove
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// // Add shape first, then remove it
    /// presentation.remove_shape(0, 0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn remove_shape(&mut self, slide_index: usize, shape_index: usize) -> Result<Shape> {
        if slide_index < self.slides.len() {
            let slide = &mut self.slides[slide_index];
            if shape_index < slide.shapes.len() {
                Ok(slide.shapes.remove(shape_index))
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Shape index {} out of bounds",
                    shape_index
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Clear all content (text and shapes) from a slide.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// presentation.clear_slide(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn clear_slide(&mut self, slide_index: usize) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].text.clear();
            self.slides[slide_index].shapes.clear();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {} out of bounds",
                slide_index
            )))
        }
    }

    /// Locate the verbatim source markup for a slide that has not been edited.
    ///
    /// Returns `None` whenever the model cannot prove the slide is byte-identical
    /// in meaning to one of the source slides, in which case the caller
    /// regenerates it from the model.
    fn retained_page(&self, slide_index: usize, slide: &Slide) -> Option<&str> {
        let source = self.content_source.as_ref()?;
        if self.declarations != self.source_declarations
            || self.page_metadata != self.source_page_metadata
        {
            return None;
        }
        if self
            .source_slides
            .get(slide_index)
            .is_some_and(|candidate| slide_content_eq(candidate, slide))
        {
            return source.page(slide_index);
        }
        if self.source_slides.len() > MAX_SOURCE_PAGE_SCAN {
            return None;
        }
        let matched = self
            .source_slides
            .iter()
            .position(|candidate| slide_content_eq(candidate, slide))?;
        source.page(matched)
    }

    /// Build the `office:automatic-styles` payload for a retained document.
    ///
    /// Only slides that were regenerated need synthesised drawing-page styles,
    /// and any name the source already defines is left untouched so retained
    /// markup keeps referring to the original definition.
    fn generate_page_styles(&self, regenerated: &[usize]) -> String {
        let defines = |name: &str| {
            self.content_source
                .as_ref()
                .is_some_and(|source| source.defines_style(name))
        };
        let mut output = String::new();
        let needs_default = regenerated.iter().any(|index| {
            self.slides.get(*index).is_some_and(|slide| {
                super::builder::slide_style_name(slide, *index)
                    == super::builder::DEFAULT_DRAWING_PAGE_STYLE_NAME
            })
        });
        if needs_default && !defines(super::builder::DEFAULT_DRAWING_PAGE_STYLE_NAME) {
            output.push_str(super::builder::DEFAULT_DRAWING_PAGE_STYLE);
        }
        for &index in regenerated {
            let Some(slide) = self.slides.get(index) else {
                continue;
            };
            if defines(&super::builder::slide_style_name(slide, index)) {
                continue;
            }
            super::builder::push_transition_style(&mut output, slide, index);
        }
        output
    }

    /// Generate content.xml from the current mutable state.
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

        let shape_count = self.slides.iter().map(|s| s.shapes.len()).sum::<usize>();
        let mut estimated = 256usize;
        estimated += self.slides.len() * 128;
        estimated += shape_count * 192;
        estimated += self
            .slides
            .iter()
            .map(|s| s.text.len() + s.title.as_ref().map(|t| t.len()).unwrap_or(0))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|sh| sh.text.len() + sh.name.as_ref().map(|n| n.len()).unwrap_or(0))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        body.push_str(&super::declaration::write_declaration_elements(
            self.declarations.as_ref(),
            self.slides.len(),
        )?);
        if let Some(source) = &self.content_source {
            source.write_leading_extras(&mut body);
        }

        let page_names = super::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        super::settings::validate_page_references(
            self.settings.as_ref(),
            &page_names,
        )?;

        let mut regenerated = Vec::new();
        for (i, slide) in self.slides.iter().enumerate() {
            if let Some(page) = self.retained_page(i, slide) {
                body.push_str(page);
                continue;
            }
            regenerated.push(i);
            let page_num = i + 1;
            let slide_style = super::builder::slide_style_name(slide, i);
            if let Some(metadata) = &self.page_metadata {
                metadata.validate_for_slides(self.slides.len())?;
            }
            let page_attributes = super::page_metadata::write_page_attributes(
                self.page_metadata.as_ref(),
                i,
                &slide_style,
            )?;
            let declaration_attributes = super::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::DeclarationTarget::Slide,
            );
            let _ = page_num;
            body.push_str("<draw:page");
            body.push_str(&page_attributes);
            body.push_str(&declaration_attributes);
            body.push('>');

            // Add title frame if title exists
            if let Some(ref title) = slide.title {
                let title_paragraphs = super::builder::generate_text_paragraphs(title, Some("P1"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr1" draw:text-style-name="P1" draw:layer="layout" presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    title_paragraphs
                ));
            }

            // Add text frame
            if !slide.text.is_empty() {
                let y_position = if slide.title.is_some() {
                    "5.0cm"
                } else {
                    "2.0cm"
                };
                let text_paragraphs =
                    super::builder::generate_text_paragraphs(&slide.text, Some("P2"));
                body.push_str(&xml_minifier::minified_xml_format!(
                    r#"<draw:frame draw:style-name="gr2" draw:text-style-name="P2" draw:layer="layout" presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y="{}"><draw:text-box>{}</draw:text-box></draw:frame>"#,
                    y_position,
                    text_paragraphs
                ));
            }

            // Add shapes
            for (shape_idx, shape) in slide.shapes.iter().enumerate() {
                body.push_str(&super::builder::Builder::generate_shape_xml(
                    shape, shape_idx,
                )?);
            }

            for animation in &slide.animations {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }
            if let Some(animation) = &slide.legacy_animation {
                animation.write_xml(&mut body, &extension_namespaces)?;
            }

            let notes_attributes = super::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::DeclarationTarget::Notes,
            );
            body.push_str(&super::declaration::apply_notes_binding(
                super::builder::Builder::generate_notes_xml(slide.notes.as_deref()),
                &notes_attributes,
            )?);

            body.push_str("</draw:page>");
        }

        if let Some(source) = &self.content_source {
            source.write_trailing_extras(&mut body);
        }
        body.push_str(&super::settings::write(
            self.settings.as_ref(),
        )?);

        if let Some(source) = &self.content_source {
            let styles = self.generate_page_styles(&regenerated);
            let synthesised = !regenerated.is_empty() || !styles.is_empty();
            let mut output = String::with_capacity(body.len() + estimated);
            source.write_prolog(&mut output);
            output.push_str(&source.root_start_tag(&extension_declarations, synthesised)?);
            source.write_prologue(&mut output, &styles);
            output.push_str(source.body_start_tag());
            output.push_str(source.presentation_start_tag());
            output.push_str(&body);
            output.push_str(&source.close_tags());
            return Ok(output);
        }

        let transition_styles = super::builder::generate_transition_styles(&self.slides);
        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"{} office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{}</office:automatic-styles><office:body><office:presentation>{}</office:presentation></office:body></office:document-content>"#,
            extension_declarations, transition_styles, body
        ))
    }

    /// Generate meta.xml with current metadata.
    fn generate_meta_xml(&self) -> Result<String> {
        if let Some(patched) = self.patched_source_meta_xml()? {
            return Ok(patched);
        }
        Ok(self.generate_meta_xml_from_scratch())
    }

    /// Patch the retained source meta.xml so metadata the edit did not change
    /// survives the save, while fields set through the mutable API, the
    /// generator, and the modification date are updated in place.
    fn patched_source_meta_xml(&self) -> Result<Option<String>> {
        let Some(package) = &self.source_package else {
            return Ok(None);
        };
        let Ok(bytes) = package.get_file("meta.xml") else {
            return Ok(None);
        };
        let Ok(source) = String::from_utf8(bytes) else {
            return Ok(None);
        };
        let source_metadata = crate::Metadata::from_xml(&source)?;
        let patch = MetaXmlPatch::preserve_all()
            .with_generator_and_modification_date("Litchi/0.0.1", chrono::Utc::now().to_rfc3339())
            .diff_simple_fields(&source_metadata, &self.metadata);
        patch_meta_xml(&source, &patch)
    }

    /// Generate meta.xml from the mutable metadata model alone.
    fn generate_meta_xml_from_scratch(&self) -> String {
        let now = chrono::Utc::now().to_rfc3339();
        let mut estimated = 64usize;
        estimated += self.metadata.title.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.author.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self.metadata.subject.as_ref().map(|s| s.len()).unwrap_or(0);
        estimated += self
            .metadata
            .description
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        estimated += self
            .metadata
            .keywords
            .as_ref()
            .map(|s| s.len())
            .unwrap_or(0);
        let mut meta_fields = String::with_capacity(estimated);

        // Add optional metadata fields
        if let Some(ref title) = self.metadata.title {
            let escaped_title = escape_xml(title);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:title>{}</dc:title>"#,
                escaped_title
            ));
        }

        if let Some(ref author) = self.metadata.author {
            let escaped_author = escape_xml(author);
            meta_fields.push_str(&xml_minifier::minified_xml_format!(
                r#"<dc:creator>{}</dc:creator>"#,
                escaped_author
            ));
        }

        xml_minifier::minified_xml_format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator><dc:date>{}</dc:date>{}</office:meta></office:document-meta>"#,
            now,
            meta_fields
        )
    }

    /// Save the modified presentation to a file.
    ///
    /// # Arguments
    ///
    /// * `path` - Path where the ODP file should be saved
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// presentation.save("output.odp")?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        let bytes = self.to_bytes()?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Convert the presentation to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odp::MutablePresentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();

        // Set MIME type
        writer.set_mimetype(&self.mimetype)?;

        // Add content.xml (regenerated from mutable state)
        let content_xml = self.generate_content_xml()?;
        writer.add_file("content.xml", content_xml.as_bytes())?;

        // Add styles.xml (preserved or default)
        let default_styles = Structure::default_styles_xml();
        let styles_xml = self.styles_xml.as_deref().unwrap_or(&default_styles);
        writer.add_file("styles.xml", styles_xml.as_bytes())?;

        // Add meta.xml (patched from the source or regenerated with current metadata)
        let meta_xml = self.generate_meta_xml()?;
        writer.add_file("meta.xml", meta_xml.as_bytes())?;

        for (path, media) in &self.media_files {
            writer.add_file_with_media_type(path, &media.bytes, &media.media_type)?;
        }

        if let Some(package) = &self.source_package {
            writer.copy_auxiliary_files_from(package)?;
        }

        writer.finish_to_bytes()
    }
}

impl Default for MutablePresentation {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        AnimationAttribute, AnimationAttributeNamespace, AnimationKind, AnimationNode,
        DrawingHyperlink, LegacyAnimationKind, LegacyAnimationNode, Action,
        Builder, EventListener, ScriptEventListener, ShapeEventListener,
    };

    const STYLES: &str = r#"<?xml version="1.0"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles><office:marker>preserve-me</office:marker></office:styles></office:document-styles>"#;
    const SETTINGS: &[u8] = b"<settings>presentation-settings</settings>";
    const IMAGE: &[u8] = b"\x89PNG\r\n\x1a\nimage-payload";
    const CUSTOM: &[u8] = b"custom-presentation-data";

    fn presentation_bytes_with_image() -> Vec<u8> {
        let content = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Media"><draw:frame presentation:class="title"><draw:text-box><text:p>Visible Title</text:p></draw:text-box></draw:frame><draw:frame presentation:class="object"><draw:text-box><text:p>Body &amp; more</text:p></draw:text-box></draw:frame><draw:frame draw:name="Photo" draw:layer="controls" draw:z-index="184467440737095516160" draw:transform="rotate (0.5)" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><draw:image xlink:href="Pictures/a&amp;b.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame><draw:rect draw:name="Labeled"><text:p>Shape label</text:p></draw:rect><draw:connector draw:name="Link" svg:x1="0cm" svg:y1="0cm" svg:x2="2cm" svg:y2="2cm"/><draw:line draw:name="Rule" svg:x1="1cm" svg:y1="1cm" svg:x2="5cm" svg:y2="1cm"/><draw:ellipse draw:name="Arc" draw:kind="section" draw:start-angle="15" draw:end-angle="275" svg:cx="3cm" svg:cy="4cm" svg:rx="2cm" svg:ry="1cm"/><draw:path draw:name="Route" svg:viewBox="0 0 100 100" svg:d="M 0 0 L 100 100"/><draw:g draw:name="Outer"><draw:rect draw:name="Grouped"><text:p>Grouped text</text:p></draw:rect><draw:g draw:name="Inner"><draw:ellipse draw:name="Nested arc" draw:kind="arc"/></draw:g></draw:g><presentation:notes><draw:frame><draw:text-box><text:p>Speaker note</text:p></draw:text-box></draw:frame></presentation:notes></draw:page></office:presentation></office:body></office:document-content>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
        writer.add_file("settings.xml", SETTINGS).unwrap();
        writer.add_file("Pictures/a&b.png", IMAGE).unwrap();
        writer
            .add_manifest_entry("Object 1/", "application/vnd.oasis.opendocument.text")
            .unwrap();
        writer
            .add_file_with_media_type("custom/data.bin", CUSTOM, "application/x-odp-test")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn mutable_presentation_round_trips_images_styles_and_settings() {
        let source_bytes = presentation_bytes_with_image();
        let presentation = Presentation::from_bytes(source_bytes.clone()).unwrap();
        assert_eq!(presentation.to_bytes().unwrap(), source_bytes);
        let source_shapes = presentation.slides().unwrap().remove(0).shapes;
        assert_eq!(source_shapes[0].image_href(), Some("Pictures/a&b.png"));

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();

        assert_eq!(package.get_file("Pictures/a&b.png").unwrap(), IMAGE);
        assert_eq!(package.get_file("settings.xml").unwrap(), SETTINGS);
        assert_eq!(package.get_file("styles.xml").unwrap(), STYLES.as_bytes());
        assert_eq!(package.get_file("custom/data.bin").unwrap(), CUSTOM);
        let borrowed = package.package().unwrap();
        assert_eq!(
            borrowed.manifest().get_media_type("Object 1/"),
            Some("application/vnd.oasis.opendocument.text")
        );
        assert_eq!(
            borrowed.manifest().get_media_type("custom/data.bin"),
            Some("application/x-odp-test")
        );

        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains("<draw:image"));
        assert!(content.contains(r#"xlink:href="Pictures/a&amp;b.png""#));
        assert!(content.contains(r#"xlink:show="embed""#));
        assert!(content.contains(r#"draw:layer="controls""#));
        assert!(content.contains(r#"draw:z-index="184467440737095516160""#));
        assert!(content.contains(r#"draw:transform="rotate (0.5)""#));
        assert!(content.contains("<draw:line"));
        assert!(content.contains("<draw:rect"));
        assert!(content.contains("<draw:connector"));
        assert!(content.contains("<draw:ellipse"));
        assert!(content.contains(r#"draw:kind="section""#));
        assert!(content.contains(r#"svg:cx="3cm""#));
        assert!(content.contains("<draw:path"));
        assert!(content.contains(r#"svg:viewBox="0 0 100 100""#));
        assert!(content.contains(r#"svg:d="M 0 0 L 100 100""#));
        assert_eq!(content.matches("<draw:g").count(), 2);
        assert!(content.contains("Grouped text"));
        assert_eq!(content.matches("Visible Title").count(), 1);
        assert_eq!(content.matches("Body &amp; more").count(), 1);
        assert_eq!(content.matches("Shape label").count(), 1);
        assert_eq!(content.matches("Speaker note").count(), 1);

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let slides = reparsed.slides().unwrap();
        assert_eq!(slides[0].title.as_deref(), Some("Visible Title"));
        assert_eq!(slides[0].text, "Body & more");
        assert_eq!(slides[0].notes.as_deref(), Some("Speaker note"));
        assert_eq!(
            slides[0].all_text(),
            "Visible Title\nBody & more\nShape label\nGrouped text"
        );
        let picture = slides[0]
            .shapes
            .iter()
            .find(|shape| shape.shape_type == litchi_core::ShapeType::Picture)
            .unwrap();
        assert_eq!(picture.image_href(), Some("Pictures/a&b.png"));
        assert_eq!(picture.layer(), Some("controls"));
        assert_eq!(picture.z_index(), Some("184467440737095516160"));
        assert_eq!(picture.transform(), Some("rotate (0.5)"));
        let arc = slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name() == Some("Arc"))
            .unwrap();
        assert_eq!(
            arc.drawing_kind(),
            Some(crate::DrawingShapeKind::Ellipse)
        );
        assert!(arc
            .drawing_attributes()
            .iter()
            .any(|attribute| attribute.local_name() == "kind" && attribute.value() == "section"));
        let group = slides[0]
            .shapes
            .iter()
            .find(|shape| shape.name() == Some("Outer"))
            .unwrap();
        assert_eq!(group.children().len(), 2);
        assert_eq!(group.children()[1].children().len(), 1);
    }

    #[test]
    fn mutable_presentation_preserves_inert_custom_shape_geometry() {
        let content = br#"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
            xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
          <o:body><o:presentation><d:page>
            <d:custom-shape d:name="Gear" r:transform="rotatex(0.25)">
              <d:enhanced-geometry d:type="gear" s:viewBox="0 0 21600 21600"
                d:enhanced-path="M 0 0 L ?f0 21600 Z" r:projection="parallel">
                <d:equation d:name="f0" d:formula="$0 + 1200"/>
                <d:handle d:handle-position="$0 10800"/>
              </d:enhanced-geometry>
            </d:custom-shape>
          </d:page></o:presentation></o:body>
        </o:document-content>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.add_file("content.xml", content).unwrap();
        let presentation = Presentation::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let regenerated = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        // The slide was never edited, so it is re-emitted byte for byte,
        // keeping the source's namespace prefixes.
        assert!(regenerated.contains("<d:enhanced-geometry"));
        assert!(regenerated.contains(r#"d:enhanced-path="M 0 0 L ?f0 21600 Z""#));
        assert!(regenerated.contains(r#"r:projection="parallel""#));
        assert!(regenerated.contains(r#"d:formula="$0 + 1200""#));
        assert!(regenerated.contains(r#"d:handle-position="$0 10800""#));

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let slides = reparsed.slides().unwrap();
        let geometry = slides[0].shapes[0].enhanced_geometry().unwrap();
        assert_eq!(geometry.children().len(), 2);
    }

    #[test]
    fn mutable_presentation_preserves_recursive_inert_three_dimensional_scenes() {
        let content = br##"<o:document-content
            xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
            xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
          <o:body><o:presentation><d:page>
            <r:scene s:x="1cm" s:width="8cm" r:projection="perspective" r:ambient-color="#102030">
              <r:light r:direction="(0 0 -1)" r:enabled="true"/>
              <r:cube r:min-edge="(-1 -1 -1)" r:max-edge="(1 1 1)"/>
              <r:scene r:shade-mode="gouraud"><r:sphere r:size="(2 2 2)"/></r:scene>
            </r:scene>
          </d:page></o:presentation></o:body>
        </o:document-content>"##;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.add_file("content.xml", content).unwrap();
        let presentation = Presentation::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let regenerated = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        // The slide was never edited, so it is re-emitted byte for byte,
        // keeping the source's namespace prefixes.
        assert!(regenerated.contains("<r:scene"));
        assert!(regenerated.contains(r##"r:ambient-color="#102030""##));
        assert!(regenerated.contains(r#"r:min-edge="(-1 -1 -1)""#));
        assert!(regenerated.contains(r#"r:shade-mode="gouraud""#));

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let scene = &reparsed.slides().unwrap()[0].shapes[0];
        assert_eq!(
            scene.drawing_kind(),
            Some(crate::DrawingShapeKind::ThreeDimensionalScene)
        );
        assert_eq!(scene.children().len(), 3);
        assert_eq!(scene.children()[2].children().len(), 1);
    }

    #[test]
    fn builder_and_mutable_presentation_round_trip_animation_trees() {
        let mut parameter = AnimationNode::new(AnimationKind::Parameter);
        parameter.set_attribute(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Animation,
                "name",
                "destination",
            )
            .unwrap(),
        );
        parameter.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Animation, "value", "2 & next")
                .unwrap(),
        );
        let mut command = AnimationNode::new(AnimationKind::Command);
        command.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Animation, "command", "show")
                .unwrap(),
        );
        command.add_child(parameter).unwrap();

        let mut root = AnimationNode::new(AnimationKind::Sequence);
        root.set_attribute(
            AnimationAttribute::new(AnimationAttributeNamespace::Smil, "begin", "slide.begin")
                .unwrap(),
        );
        root.set_attribute(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Other("urn:example:timing".to_string()),
                "mode",
                "author-defined",
            )
            .unwrap(),
        );
        root.add_child(command).unwrap();
        root.add_child(AnimationNode::new(AnimationKind::TransitionFilter))
            .unwrap();

        let slide = Slide {
            title: Some("Animated".to_string()),
            text: String::new(),
            index: 0,
            notes: None,
            transition: None,
            animations: vec![root.clone()],
            legacy_animation: None,
            shapes: Vec::new(),
        };
        let mut builder = Builder::new();
        builder.add_slide_element(slide).unwrap();
        let built = builder.build().unwrap();
        let presentation = Presentation::from_bytes(built).unwrap();
        assert_eq!(presentation.slides().unwrap()[0].animations, [root.clone()]);

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(content.contains(r#"xmlns:anim-ext1="urn:example:timing""#));
        assert!(content.contains(r#"anim-ext1:mode="author-defined""#));
        assert!(content.contains(r#"anim:value="2 &amp; next""#));

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        assert_eq!(reparsed.slides().unwrap()[0].animations, [root]);
    }

    #[test]
    fn rejects_invalid_mutated_animation_trees_and_xml_characters() {
        let mut leaf = AnimationNode::new(AnimationKind::Animate);
        leaf.children_mut()
            .push(AnimationNode::new(AnimationKind::Set));
        let mut presentation = MutablePresentation::new();
        presentation.add_slide("Invalid", "").unwrap();
        presentation.slides_mut()[0].animations.push(leaf);
        assert!(presentation.to_bytes().is_err());

        assert!(
            AnimationAttribute::new(AnimationAttributeNamespace::Smil, "begin", "bad\0value")
                .is_err()
        );
        assert!(
            AnimationAttribute::new(
                AnimationAttributeNamespace::Other(
                    "http://www.w3.org/XML/1998/namespace".to_string()
                ),
                "id",
                "bad namespace variant"
            )
            .is_err()
        );
    }

    #[test]
    fn mutable_presentation_preserves_and_adds_embedded_media() {
        const ORIGINAL: &[u8] = b"original-video";
        const ADDED: &[u8] = b"added-audio";
        let mut builder = Builder::new();
        let original = builder
            .embed_media("Media/original.mp4", ORIGINAL, "video/mp4")
            .unwrap();
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: None,
                shapes: vec![Shape::new().with_media(original)],
            })
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();
        assert!(
            mutable
                .embed_media("Media/original.mp4", b"replacement", "video/mp4")
                .is_err()
        );
        let added = mutable
            .embed_media("Media/added.ogg", ADDED, "audio/ogg")
            .unwrap();
        mutable
            .add_shape(0, Shape::new().with_media(added))
            .unwrap();

        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        assert_eq!(package.get_file("Media/original.mp4").unwrap(), ORIGINAL);
        assert_eq!(package.get_file("Media/added.ogg").unwrap(), ADDED);
        assert_eq!(
            package
                .package()
                .unwrap()
                .manifest()
                .get_media_type("Media/added.ogg"),
            Some("audio/ogg")
        );

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let slides = reparsed.slides().unwrap();
        assert_eq!(slides[0].shapes.len(), 2);
        assert_eq!(
            slides[0].shapes[0].media().unwrap().href(),
            "Media/original.mp4"
        );
        assert_eq!(
            slides[0].shapes[1].media().unwrap().href(),
            "Media/added.ogg"
        );
    }

    #[test]
    fn builder_and_mutable_round_trip_legacy_presentation_effects() {
        let attr =
            |namespace, name, value| AnimationAttribute::new(namespace, name, value).unwrap();
        let mut sound = LegacyAnimationNode::new(LegacyAnimationKind::Sound);
        sound.set_attribute(attr(
            AnimationAttributeNamespace::Xlink,
            "href",
            "Sounds/chime.ogg",
        ));
        sound.set_attribute(attr(AnimationAttributeNamespace::Xlink, "type", "simple"));
        let mut show = LegacyAnimationNode::new(LegacyAnimationKind::ShowShape);
        show.set_attribute(attr(
            AnimationAttributeNamespace::Draw,
            "shape-id",
            "shape1",
        ));
        show.set_attribute(attr(
            AnimationAttributeNamespace::Presentation,
            "effect",
            "fade",
        ));
        show.add_child(sound).unwrap();
        let mut root = LegacyAnimationNode::new(LegacyAnimationKind::Animations);
        root.set_attribute(attr(
            AnimationAttributeNamespace::Other("urn:example:legacy-effects".to_string()),
            "mode",
            "preserve",
        ));
        root.add_child(show).unwrap();

        let mut builder = Builder::new();
        builder
            .add_slide_element(Slide {
                title: None,
                text: String::new(),
                index: 0,
                notes: None,
                transition: None,
                animations: Vec::new(),
                legacy_animation: Some(root.clone()),
                shapes: Vec::new(),
            })
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        assert_eq!(
            presentation.slides().unwrap()[0].legacy_animation(),
            Some(&root)
        );

        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let reparsed = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        assert_eq!(
            reparsed.slides().unwrap()[0].legacy_animation(),
            Some(&root)
        );
    }

    #[test]
    fn mutable_presentation_preserves_shape_links_and_inert_actions() {
        let mut shape = Shape::new();
        shape.set_hyperlink(Some(DrawingHyperlink::new("#page2").unwrap()));
        shape
            .add_event_listener(ShapeEventListener::Script(
                ScriptEventListener::external_binding(
                    "dom:mouseover",
                    "javascript",
                    "Scripts/hover.js",
                )
                .unwrap(),
            ))
            .unwrap();
        shape
            .add_event_listener(ShapeEventListener::Action(Box::new(
                EventListener::new("dom:click", Action::NextPage).unwrap(),
            )))
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
                shapes: vec![shape.clone()],
            })
            .unwrap();
        let presentation = Presentation::from_bytes(builder.build().unwrap()).unwrap();
        let mutable = MutablePresentation::from_presentation(presentation).unwrap();
        let reparsed = Presentation::from_bytes(mutable.to_bytes().unwrap()).unwrap();
        let parsed = &reparsed.slides().unwrap()[0].shapes[0];
        assert_eq!(parsed.hyperlink(), shape.hyperlink());
        assert_eq!(parsed.event_listeners(), shape.event_listeners());
    }

    #[test]
    fn regenerated_media_frame_keeps_its_fallback_image() {
        let content = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="page1"><draw:frame draw:name="Movie"><draw:plugin xlink:href="Models/duck.json" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="model/vnd.gltf+json"/><draw:image xlink:href="Pictures/fallback.png" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;
        let mut writer = PackageWriter::new();
        writer
            .set_mimetype("application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.add_file("content.xml", content.as_bytes()).unwrap();
        let presentation = Presentation::from_bytes(writer.finish_to_bytes().unwrap()).unwrap();
        let mut mutable = MutablePresentation::from_presentation(presentation).unwrap();

        // Editing the slide forces the page to be rebuilt from the model, which
        // is where a media frame used to be rejected outright.
        mutable.slides_mut()[0].text = "Edited".to_string();
        let bytes = mutable.to_bytes().unwrap();
        let package = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let regenerated = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
        assert!(regenerated.contains(r#"xlink:href="Models/duck.json""#));
        assert!(regenerated.contains(r#"xlink:href="Pictures/fallback.png""#));

        let reparsed = Presentation::from_bytes(bytes).unwrap();
        let shapes = reparsed.slides().unwrap().remove(0).shapes;
        let frame = shapes
            .iter()
            .find(|shape| shape.media().is_some())
            .expect("media frame");
        assert_eq!(frame.image_href(), Some("Pictures/fallback.png"));
    }
}
