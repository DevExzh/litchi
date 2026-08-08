//! Private staging engine for source-checked presentation transactions.
//!
//! The attached mutable root is intentionally not exported. Public mutation is
//! available only through [`super::edit::Transaction`].

use crate::codec::content_source::ContentSource;
use crate::core::{OwnedPackage, PackageWriter, Structure};
use crate::model::animation::validate_animation_roots;
use crate::model::legacy_animation::validate_legacy_animation_root;
use crate::model::media::{EmbeddedMedia, embed_media, validate_package_media_path};
use crate::{Presentation, Reference, Shape, Slide};
use litchi_core::{Result, xml::escape_xml};
use std::collections::{BTreeMap, BTreeSet};

/// A mutable ODP presentation that supports in-place modifications.
///
/// This struct wraps an ODP presentation and provides methods to modify its content,
/// including adding, updating, and removing slides and shapes.
///
/// # Examples
///
/// ```no_run
/// use litchi_odp::Presentation;
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
pub(super) struct MutablePresentation {
    /// Mutable slides
    slides: Vec<Slide>,
    /// Original MIME type
    mimetype: String,
    /// Original styles XML (preserved as-is)
    styles_xml: Option<String>,
    /// Original package retained for copying auxiliary package parts.
    source_package: Option<OwnedPackage>,
    /// Newly embedded package media, keyed by package path.
    media_files: BTreeMap<String, EmbeddedMedia>,
    /// Inert slide-show settings and custom shows.
    settings: Option<crate::Settings>,
    /// Inert header/footer/date-time declarations and page bindings.
    declarations: Option<crate::model::declaration::Collection>,
    /// Static page names, IDs, and layout/master references.
    page_metadata: Option<crate::model::page_metadata::Collection>,
    /// Verbatim fragments of the source `content.xml`, when one was opened.
    ///
    /// Slides the caller never touched are re-emitted from these fragments so
    /// constructs outside the model — nested tables, image text alternatives,
    /// automatic styles, font declarations — survive a save.
    content_source: Option<ContentSource>,
    /// Slides exactly as parsed, used to detect which pages are still pristine.
    source_slides: Vec<Slide>,
    /// Exact source-page identity for every staged slide; inserted slides have no source origin.
    origins: Vec<Option<usize>>,
    /// Declaration state exactly as parsed; changing it rewrites every page.
    source_declarations: Option<crate::model::declaration::Collection>,
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
    /// use litchi_odp::Presentation;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let presentation = Presentation::open("slides.odp")?;
    /// let mut mutable = MutablePresentation::from_presentation(presentation)?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn from_presentation(presentation: &Presentation) -> Result<Self> {
        let slides = presentation.slides()?;
        let settings = presentation.settings()?;
        let parsed_declarations = presentation.declarations()?;
        let parsed_page_metadata = presentation.pages()?;
        let mimetype = presentation.owned_package().mimetype()?;

        let styles_xml = presentation.styles_xml().map(str::to_owned);
        let content_source = ContentSource::parse(presentation.content_xml())?;
        let source_package = Some(presentation.owned_package().clone());

        let declarations = (!parsed_declarations.is_empty()).then_some(parsed_declarations);
        let page_metadata = (!parsed_page_metadata.is_empty()).then_some(parsed_page_metadata);
        let source_slides = match &content_source {
            // Only pages the scanner actually recovered can be retained; a
            // mismatch means the model and the markup disagree, so fall back to
            // regenerating every page.
            Some(source) if source.page_count() == slides.len() => slides.clone(),
            _ if slides.is_empty() => Vec::new(),
            _ => {
                return Err(litchi_core::Error::Unsupported(
                    "ODP source slide coverage is incomplete; transaction staging refuses regeneration"
                        .to_string(),
                ));
            },
        };
        let retained_source = content_source.filter(|_| !source_slides.is_empty());
        let origins = if retained_source.is_some() {
            (0..slides.len()).map(Some).collect()
        } else {
            vec![None; slides.len()]
        };

        Ok(Self {
            slides,
            mimetype,
            styles_xml,
            source_package,
            media_files: BTreeMap::new(),
            settings,
            source_declarations: declarations.clone(),
            declarations,
            page_metadata,
            content_source: retained_source,
            source_slides,
            origins,
        })
    }

    /// Get all slides in the presentation.
    pub(super) fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Return whether rewriting this slide would discard retained source XML.
    pub(super) fn retains_source_slide(&self, index: usize) -> bool {
        self.slides
            .get(index)
            .is_some_and(|slide| self.retained_page(index, slide).is_some())
    }

    /// Return whether structural edits would need declaration rebinding.
    pub(super) fn has_source_declarations(&self) -> bool {
        self.source_declarations.is_some()
    }

    /// Add a package-contained audio or video payload.
    ///
    /// Existing source-package paths cannot be replaced implicitly. The
    /// returned inert reference can be attached with [`Shape::with_media`].
    pub(super) fn embed_media(
        &mut self,
        path: String,
        bytes: Vec<u8>,
        media_type: String,
    ) -> Result<Reference> {
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

    /// Verify that every media part staged by this transaction survived package publication.
    ///
    /// This deliberately verifies package bytes and manifest metadata independently of slide
    /// references: callers may stage media before attaching it to a drawing frame.
    pub(super) fn verify_embedded_media(&self, reopened: &OwnedPackage) -> Result<()> {
        let package = reopened.package()?;
        for (path, expected) in &self.media_files {
            if !package.has_file(path) {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP package is missing staged media path '{path}'"
                )));
            }
            let actual = package.get_file(path)?;
            if actual != expected.bytes {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP media bytes differ for '{path}'"
                )));
            }
            let Some(entry) = package.manifest().get_entry(path) else {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP manifest is missing staged media entry '{path}'"
                )));
            };
            if entry.media_type != expected.media_type {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "published ODP media type differs for '{path}'"
                )));
            }
        }
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
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("First", "Content 1")?;
    /// presentation.add_slide("Third", "Content 3")?;
    /// presentation.insert_slide(1, "Second", "Content 2")?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn insert_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
        if index <= self.slides.len() {
            let candidate_metadata = crate::model::page_metadata::metadata_after_page_insert(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = crate::model::page_metadata::effective_page_names(
                Some(&candidate_metadata),
                self.slides.len() + 1,
            )?;
            crate::model::settings::validate_page_references(
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
            self.origins.insert(index, None);
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
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content 1")?;
    /// presentation.add_slide("Slide 2", "Content 2")?;
    /// presentation.remove_slide(0)?; // Remove first slide
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn remove_slide(&mut self, index: usize) -> Result<Slide> {
        if index < self.slides.len() {
            let current_names = crate::model::page_metadata::effective_page_names(
                self.page_metadata.as_ref(),
                self.slides.len(),
            )?;
            let removed_name = &current_names[index];
            if crate::model::settings::settings_reference_page(self.settings.as_ref(), removed_name)
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "cannot remove presentation page '{removed_name}' because slide-show settings reference it"
                )));
            }
            let candidate_metadata = crate::model::page_metadata::metadata_after_page_remove(
                self.page_metadata.as_ref(),
                self.slides.len(),
                index,
            )?;
            let candidate_names = crate::model::page_metadata::effective_page_names(
                candidate_metadata.as_ref(),
                self.slides.len() - 1,
            )?;
            crate::model::settings::validate_page_references(
                self.settings.as_ref(),
                &candidate_names,
            )?;
            let slide = self.slides.remove(index);
            self.origins.remove(index);
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
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Old Title", "Old content")?;
    /// presentation.update_slide(0, "New Title", "New content")?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn update_slide(&mut self, index: usize, title: &str, text: &str) -> Result<()> {
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
    /// use litchi_odp::Shape;
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
    pub(super) fn add_shape(&mut self, slide_index: usize, shape: Shape) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].shapes.push(shape);
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )))
        }
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
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// // Add shape first, then remove it
    /// presentation.remove_shape(0, 0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn remove_shape(&mut self, slide_index: usize, shape_index: usize) -> Result<Shape> {
        if slide_index < self.slides.len() {
            let slide = &mut self.slides[slide_index];
            if shape_index < slide.shapes.len() {
                Ok(slide.shapes.remove(shape_index))
            } else {
                Err(litchi_core::Error::InvalidFormat(format!(
                    "Shape index {shape_index} out of bounds"
                )))
            }
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
            )))
        }
    }

    /// Clear all content (text and shapes) from a slide.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// presentation.clear_slide(0)?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn clear_slide(&mut self, slide_index: usize) -> Result<()> {
        if slide_index < self.slides.len() {
            self.slides[slide_index].title = None;
            self.slides[slide_index].text.clear();
            self.slides[slide_index].shapes.clear();
            Ok(())
        } else {
            Err(litchi_core::Error::InvalidFormat(format!(
                "Slide index {slide_index} out of bounds"
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
        if self.declarations != self.source_declarations {
            return None;
        }
        let source_index = self.origins.get(slide_index).copied().flatten()?;
        self.source_slides
            .get(source_index)
            .filter(|candidate| slide_content_eq(candidate, slide))?;
        source.page(source_index)
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
            .map(|slide| slide.text.len() + slide.title.as_ref().map_or(0, String::len))
            .sum::<usize>();
        estimated += self
            .slides
            .iter()
            .flat_map(|s| s.shapes.iter())
            .map(|shape| shape.text.len() + shape.name.as_ref().map_or(0, String::len))
            .sum::<usize>();

        let mut body = String::with_capacity(estimated);

        body.push_str(&crate::model::declaration::write_declaration_elements(
            self.declarations.as_ref(),
            self.slides.len(),
        )?);
        if let Some(source) = &self.content_source {
            source.write_leading_extras(&mut body);
        }

        let page_names = crate::model::page_metadata::effective_page_names(
            self.page_metadata.as_ref(),
            self.slides.len(),
        )?;
        crate::model::settings::validate_page_references(self.settings.as_ref(), &page_names)?;

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

            let notes_attributes = crate::model::declaration::write_binding_attributes(
                self.declarations.as_ref(),
                i,
                crate::model::declaration::Target::Notes,
            );
            body.push_str(&crate::model::declaration::apply_notes_binding(
                super::builder::Builder::generate_notes_xml(slide.notes.as_deref()),
                &notes_attributes,
            )?);

            body.push_str("</draw:page>");
        }

        if let Some(source) = &self.content_source {
            source.write_trailing_extras(&mut body);
        }
        body.push_str(&crate::model::settings::write(self.settings.as_ref())?);

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
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"{extension_declarations} office:version="1.3"><office:scripts/><office:font-face-decls/><office:automatic-styles>{transition_styles}</office:automatic-styles><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
        ))
    }

    /// Preserve source metadata exactly or create deterministic minimal metadata.
    fn generate_meta_xml(&self) -> Result<String> {
        if let Some(package) = &self.source_package
            && let Ok(bytes) = package.get_file("meta.xml")
        {
            return String::from_utf8(bytes).map_err(|error| {
                litchi_core::Error::InvalidFormat(format!(
                    "ODP source meta.xml is not UTF-8: {error}"
                ))
            });
        }
        Ok(r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>Litchi/0.0.1</meta:generator></office:meta></office:document-meta>"#.to_string())
    }

    /// Convert the presentation to bytes.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// The public equivalent starts from `Presentation::snapshot()`.
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut presentation = MutablePresentation::new();
    /// presentation.add_slide("Slide 1", "Content")?;
    /// let bytes = presentation.to_bytes()?;
    /// # Ok(())
    /// # }
    /// ```
    pub(super) fn to_bytes_bounded(&self, limit: usize) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new_bounded(limit);

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

        writer.finish_to_bounded_bytes()
    }
}

/// Compare two slides ignoring their position in the deck.
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
