/// Slide-related objects, including Slide, SlideLayout, and SlideMaster.
use crate::error::{OoxmlError, Result};
use crate::pptx::parts::{SlideLayoutPart, SlideMasterPart, SlidePart, ThemePart};
use crate::pptx::shapes::base::BaseShape;
use litchi_opc::OpcPackage;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::Part;
use quick_xml::Reader;
use quick_xml::events::Event;

const STRICT_SLIDE_LAYOUT_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout";
const STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster";
const STRICT_THEME_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/theme";

fn resolve_picture_background(
    source_part: &dyn Part,
    package: &OpcPackage,
) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
    let Some((relationship_id, style)) = picture_background_reference(source_part.blob())? else {
        return Ok(None);
    };
    let relationship = source_part.rels().get(&relationship_id).ok_or_else(|| {
        OoxmlError::InvalidRelationship(format!(
            "background image relationship '{relationship_id}' does not exist"
        ))
    })?;
    if !matches!(relationship.reltype(), rt::IMAGE | rt::STRICT_IMAGE) {
        return Err(OoxmlError::InvalidRelationship(format!(
            "background image relationship '{relationship_id}' has unsupported type '{}'",
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(format!(
            "background image relationship '{relationship_id}' must be internal"
        )));
    }

    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "invalid background image relationship '{relationship_id}': {error}"
        ))
    })?;
    let image_part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "background image relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    let image_data = image_part.blob();
    let format = crate::pptx::format::ImageFormat::detect_from_bytes(image_data).ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "background image relationship '{relationship_id}' targets an unsupported image format"
        ))
    })?;

    Ok(Some(crate::pptx::backgrounds::SlideBackground::Picture {
        image_data: image_data.to_vec(),
        format,
        style,
    }))
}

fn picture_background_reference(
    xml: &[u8],
) -> Result<Option<(String, crate::pptx::backgrounds::PictureStyle)>> {
    let xml = crate::common::mce::process_ooxml(xml)?;
    let mut reader = Reader::from_reader(xml.as_ref());
    let mut in_background = false;
    let mut in_blip_fill = false;
    let mut relationship_id = None;
    let mut style = crate::pptx::backgrounds::PictureStyle::Stretch;

    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) => match element.local_name().as_ref() {
                b"bg" => in_background = true,
                b"blipFill" if in_background => in_blip_fill = true,
                b"blip" if in_blip_fill => {
                    relationship_id = crate::drawings::blip::read_blip_embed_attr(&element)?;
                },
                b"tile" if in_blip_fill => {
                    style = crate::pptx::backgrounds::PictureStyle::Tile;
                },
                _ => {},
            },
            Ok(Event::Empty(element)) => {
                if in_blip_fill && element.local_name().as_ref() == b"blip" {
                    relationship_id = crate::drawings::blip::read_blip_embed_attr(&element)?;
                } else if in_blip_fill && element.local_name().as_ref() == b"tile" {
                    style = crate::pptx::backgrounds::PictureStyle::Tile;
                }
            },
            Ok(Event::End(element)) => {
                if in_blip_fill && element.local_name().as_ref() == b"blipFill" {
                    return Ok(relationship_id.map(|id| (id, style)));
                }
                if element.local_name().as_ref() == b"bg" {
                    in_background = false;
                }
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }

    Ok(None)
}

/// A slide in a presentation.
///
/// Provides access to slide content and properties, following the python-pptx
/// interface design.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::pptx::Package;
///
/// let pkg = Package::open("presentation.pptx")?;
/// let pres = pkg.presentation()?;
///
/// for slide in pres.slides()?.iter() {
///     println!("Slide name: {}", slide.name()?);
///     println!("Text content: {}", slide.text()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct Slide<'a> {
    /// The underlying slide part
    part: SlidePart<'a>,
    /// Reference to the OPC package (for accessing notes)
    package: Option<&'a litchi_opc::OpcPackage>,
}

#[allow(dead_code)] // Part of the public API for future use
impl<'a> Slide<'a> {
    /// Create a new Slide from a SlidePart.
    ///
    /// This is typically called internally.
    #[inline]
    pub(crate) fn new(part: SlidePart<'a>) -> Self {
        Self {
            part,
            package: None,
        }
    }

    /// Create a new Slide with a reference to the package.
    ///
    /// This allows accessing related parts like notes.
    #[inline]
    pub(crate) fn with_package(part: SlidePart<'a>, package: &'a litchi_opc::OpcPackage) -> Self {
        Self {
            part,
            package: Some(package),
        }
    }

    /// Get the slide name.
    ///
    /// Returns the internal name of the slide from the `<p:cSld>` element.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     println!("First slide name: {}", slide.name()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Extract all text content from the slide.
    ///
    /// This extracts text from all text elements in the slide,
    /// including shapes, text boxes, and tables.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    ///
    /// for slide in pres.slides()?.iter() {
    ///     let text = slide.text()?;
    ///     if !text.is_empty() {
    ///         println!("Slide content:\n{}", text);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        self.part.extract_text()
    }

    /// Get access to the underlying slide part.
    ///
    /// This provides lower-level access to the slide XML.
    #[inline]
    pub fn part(&self) -> &SlidePart<'a> {
        &self.part
    }

    /// Load all programmable tag-list parts related to this slide.
    pub fn tag_lists(&self) -> Result<Vec<crate::pptx::tags::SlideTagList>> {
        let package = self.package.ok_or_else(|| {
            crate::error::OoxmlError::InvalidFormat(
                "slide tag lists require package-backed slide access".into(),
            )
        })?;
        crate::pptx::tags::load_slide_tag_lists(self.part.part(), package)
    }

    /// Get all shapes on this slide.
    ///
    /// Returns a vector of BaseShape objects that provide access to text,
    /// pictures, tables, and other shape types.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    /// use litchi_ooxml::pptx::shapes::ShapeType;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     for shape in slide.shapes()? {
    ///         let mut shape_mut = shape;
    ///         match shape_mut.shape_type() {
    ///             ShapeType::Shape => {
    ///                 println!("Text shape: {}", shape_mut.name()?);
    ///             }
    ///             ShapeType::Picture => {
    ///                 println!("Picture: {}", shape_mut.name()?);
    ///             }
    ///             ShapeType::GraphicFrame if shape_mut.has_table() => {
    ///                 println!("Table: {}", shape_mut.name()?);
    ///             }
    ///             _ => {}
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        self.part.shapes()
    }

    /// Get all placeholder shapes on this slide.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        self.part.placeholders()
    }

    /// Get the number of shapes on this slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     println!("Shape count: {}", slide.shape_count()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shape_count(&self) -> Result<usize> {
        Ok(self.shapes()?.len())
    }

    /// Get a specific shape by index.
    ///
    /// # Arguments
    /// * `index` - Zero-based index of the shape
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if let Some(shape) = slide.shape(0)? {
    ///         let mut shape_mut = shape;
    ///         println!("First shape: {}", shape_mut.name()?);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn shape(&self, index: usize) -> Result<Option<BaseShape>> {
        Ok(self.shapes()?.into_iter().nth(index))
    }

    /// Get a shape by its non-visual shape ID.
    ///
    /// Shape IDs are used by PowerPoint animation and timing records. Returns
    /// None if the slide has no shape with this ID, and returns an error when
    /// the slide contains duplicate matching IDs.
    pub fn shape_by_id(&self, id: u32) -> Result<Option<BaseShape>> {
        let mut matched = None;
        for shape in self.shapes()? {
            if shape.shape_id()? == Some(id) {
                if matched.is_some() {
                    return Err(OoxmlError::InvalidFormat(format!(
                        "slide contains multiple shapes with non-visual ID {id}"
                    )));
                }
                matched = Some(shape);
            }
        }
        Ok(matched)
    }

    /// Check if the slide has any tables.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if slide.has_tables()? {
    ///         println!("Slide contains tables");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn has_tables(&self) -> Result<bool> {
        for shape in self.shapes()? {
            if shape.has_table() {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Check if the slide has any pictures.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if slide.has_pictures()? {
    ///         println!("Slide contains pictures");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn has_pictures(&self) -> Result<bool> {
        use crate::pptx::shapes::ShapeType;

        for shape in self.shapes()? {
            if matches!(shape.shape_type(), ShapeType::Picture) {
                return Ok(true);
            }
        }
        Ok(false)
    }

    /// Get all text shapes from this slide.
    ///
    /// Returns shapes that contain text (excluding pictures and other non-text shapes).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     for mut shape in slide.text_shapes()? {
    ///         if let Some(text) = shape.text()? {
    ///             println!("Text: {}", text);
    ///         }
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn text_shapes(&self) -> Result<Vec<BaseShape>> {
        let mut text_shapes = Vec::new();

        for shape in self.shapes()? {
            if shape.text()?.is_some() {
                text_shapes.push(shape);
            }
        }

        Ok(text_shapes)
    }

    /// Find text in the slide.
    ///
    /// Returns indices of shapes that contain the search text.
    ///
    /// # Arguments
    /// * `query` - Text to search for (case-sensitive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     let matches = slide.find_text("important")?;
    ///     println!("Found {} matching shapes", matches.len());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn find_text(&self, query: &str) -> Result<Vec<usize>> {
        let mut matches = Vec::new();

        for (idx, shape) in self.shapes()?.into_iter().enumerate() {
            if let Some(text) = shape.text()?
                && text.contains(query)
            {
                matches.push(idx);
            }
        }

        Ok(matches)
    }

    /// Check if the slide is empty (has no shapes).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// for slide in slides {
    ///     if slide.is_empty()? {
    ///         println!("Empty slide found");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn is_empty(&self) -> Result<bool> {
        Ok(self.shape_count()? == 0)
    }

    /// Get the transition effect for this slide.
    ///
    /// Returns `None` if no transition is defined for this slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if let Some(transition) = slide.transition()? {
    ///         println!("Transition type: {:?}", transition.transition_type);
    ///         println!("Speed: {:?}", transition.speed);
    ///     } else {
    ///         println!("No transition");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        self.part.transition()
    }

    /// Resolve the layout used by this slide.
    ///
    /// The layout relationship must be an internal PresentationML slide-layout
    /// part. A slide with no layout relationship or more than one layout
    /// relationship is invalid.
    pub fn layout(&self) -> Result<SlideLayout<'a>> {
        let package = self.package.ok_or_else(|| {
            OoxmlError::InvalidFormat(
                "slide-layout discovery requires package-backed slide access".to_string(),
            )
        })?;
        let slide_part = self.part.part();
        let mut layout_relationships = slide_part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_RELATIONSHIP_TYPE
            )
        });
        let relationship = layout_relationships.next().ok_or_else(|| {
            OoxmlError::InvalidRelationship(
                "slide does not have a slide-layout relationship".to_string(),
            )
        })?;
        if layout_relationships.next().is_some() {
            return Err(OoxmlError::InvalidRelationship(
                "slide has multiple slide-layout relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide-layout relationship '{}' must be internal",
                relationship.r_id()
            )));
        }

        let part_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid slide-layout relationship '{}': {error}",
                relationship.r_id()
            ))
        })?;
        let layout_part = package.get_part(&part_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "slide-layout relationship '{}' targets missing part '{}': {error}",
                relationship.r_id(),
                part_name.as_str()
            ))
        })?;
        if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::PML_SLIDE_LAYOUT.to_string(),
                got: layout_part.content_type().to_string(),
            });
        }

        Ok(SlideLayout::with_package(
            SlideLayoutPart::from_part(layout_part)?,
            package,
        ))
    }

    /// Resolve the slide master used by this slide.
    ///
    /// The master is resolved through this slide's layout. Invalid layout or
    /// master relationships are returned as errors.
    pub fn master(&self) -> Result<SlideMaster<'a>> {
        self.layout()?.master()
    }

    /// Resolve the Office theme inherited by this slide.
    ///
    /// The theme is resolved through this slide's layout and slide master.
    /// Invalid layout, master, or theme relationships are returned as errors.
    pub fn theme(&self) -> Result<crate::pptx::parts::Theme> {
        self.layout()?.theme()
    }

    /// Get the transition this slide will use after inheritance is applied.
    ///
    /// A transition defined on the slide takes precedence over its layout and
    /// slide master. If the slide has no transition, the layout's effective
    /// transition is returned instead.
    pub fn effective_transition(
        &self,
    ) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        if let Some(transition) = self.transition()? {
            return Ok(Some(transition));
        }

        self.layout()?.effective_transition()
    }

    /// Get typed simple animation timing metadata for this slide.
    ///
    /// Targets are validated against shape IDs on the current slide. Unsupported
    /// timing subtrees remain inert and are not interpreted.
    pub fn animations(&self) -> Result<crate::pptx::animations::AnimationSequence> {
        self.part.animations()
    }

    /// Get the background for this slide.
    ///
    /// Returns `None` if no background is defined directly on the slide. Use
    /// [`effective_background`](Self::effective_background) to resolve layout
    /// and master inheritance.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if let Some(bg) = slide.background()? {
    ///         println!("Slide has custom background: {:?}", bg);
    ///     } else {
    ///         println!("Using inherited background");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        if let Some(package) = self.package
            && let Some(background) = resolve_picture_background(self.part.part(), package)?
        {
            return Ok(Some(background));
        }

        self.part.background()
    }

    /// Get the background this slide will use after inheritance is applied.
    ///
    /// A background defined on the slide takes precedence over its layout and
    /// slide master. If the slide has no local background, the layout's
    /// effective background is returned instead.
    pub fn effective_background(
        &self,
    ) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        if let Some(background) = self.background()? {
            return Ok(Some(background));
        }

        self.layout()?.effective_background()
    }

    /// Get the speaker notes for this slide.
    ///
    /// Returns `None` if no notes are defined or if the package reference is not available.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// for slide in slides {
    ///     if let Some(notes) = slide.notes()? {
    ///         println!("Notes: {}", notes);
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn notes(&self) -> Result<Option<String>> {
        match self.notes_resource()? {
            Some(resource) => resource.text(),
            None => Ok(None),
        }
    }

    /// Get the validated inert raw notes-slide resource for this slide.
    pub fn notes_resource(&self) -> Result<Option<crate::pptx::notes::PptxNotesSlideResource>> {
        let Some(package) = self.package else {
            return Ok(None);
        };
        crate::pptx::notes::load_slide_notes_resource(package, self.part.part().partname())
    }

    /// Extract text from notes XML.
    fn extract_notes_text(xml: &[u8]) -> Result<Option<String>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;

        loop {
            match reader.read_event() {
                Ok(Event::Start(e)) if e.local_name().as_ref() == b"t" => {
                    in_text_element = true;
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| crate::error::OoxmlError::Xml(e.to_string()))?;
                    if !text.is_empty() && !text.ends_with('\n') {
                        text.push('\n');
                    }
                    text.push_str(t);
                },
                Ok(Event::End(e)) if e.local_name().as_ref() == b"t" => {
                    in_text_element = false;
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::error::OoxmlError::Xml(e.to_string())),
                _ => {},
            }
        }

        if text.is_empty() {
            Ok(None)
        } else {
            Ok(Some(text))
        }
    }
}

/// A slide layout.
///
/// Slide layouts define the arrangement of placeholders and other elements
/// that slides based on this layout inherit.
///
/// # Examples
///
/// ```rust,ignore
/// let layout = slide.layout()?;
/// println!("Layout name: {}", layout.name()?);
/// ```
pub struct SlideLayout<'a> {
    /// The underlying slide layout part
    part: SlideLayoutPart<'a>,
    /// Package used to resolve the layout's owning master.
    package: &'a OpcPackage,
}

impl<'a> SlideLayout<'a> {
    /// Create a new SlideLayout with its owning package.
    #[allow(unused)]
    #[inline]
    pub(crate) fn with_package(part: SlideLayoutPart<'a>, package: &'a OpcPackage) -> Self {
        Self { part, package }
    }

    /// Get the layout name.
    ///
    /// # Examples
    ///
    /// ```rust,ignore
    /// let name = layout.name()?;
    /// println!("Layout: {}", name);
    /// ```
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Get all shapes defined by this slide layout.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        self.part.shapes()
    }

    /// Get all placeholder shapes defined by this slide layout.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        self.part.placeholders()
    }

    /// Get the transition effect inherited from this layout.
    ///
    /// Returns `None` if the layout has no transition.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        self.part.transition()
    }

    /// Get the transition this layout will use after inheritance is applied.
    ///
    /// A transition defined on the layout takes precedence over its slide
    /// master. If the layout has no transition, the master's transition is
    /// returned instead.
    pub fn effective_transition(
        &self,
    ) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        if let Some(transition) = self.transition()? {
            return Ok(Some(transition));
        }

        self.master()?.transition()
    }

    /// Resolve the slide master that owns this layout.
    ///
    /// The master relationship must be internal and target a PresentationML
    /// slide-master part. A layout with no master relationship or more than one
    /// master relationship is invalid.
    pub fn master(&self) -> Result<SlideMaster<'a>> {
        let layout_part = self.part.part();
        let mut master_relationships = layout_part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::SLIDE_MASTER | STRICT_SLIDE_MASTER_RELATIONSHIP_TYPE
            )
        });
        let relationship = master_relationships.next().ok_or_else(|| {
            OoxmlError::InvalidRelationship(
                "slide layout does not have a slide-master relationship".to_string(),
            )
        })?;
        if master_relationships.next().is_some() {
            return Err(OoxmlError::InvalidRelationship(
                "slide layout has multiple slide-master relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide-master relationship '{}' must be internal",
                relationship.r_id()
            )));
        }

        let part_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid slide-master relationship '{}': {error}",
                relationship.r_id()
            ))
        })?;
        let master_part = self.package.get_part(&part_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "slide-master relationship '{}' targets missing part '{}': {error}",
                relationship.r_id(),
                part_name.as_str()
            ))
        })?;
        if master_part.content_type() != ct::PML_SLIDE_MASTER {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::PML_SLIDE_MASTER.to_string(),
                got: master_part.content_type().to_string(),
            });
        }

        Ok(SlideMaster::with_package(
            SlideMasterPart::from_part(master_part)?,
            self.package,
        ))
    }

    /// Resolve the Office theme inherited by this layout.
    ///
    /// The theme is owned by the layout's slide master. Invalid master or theme
    /// relationships are returned as errors.
    pub fn theme(&self) -> Result<crate::pptx::parts::Theme> {
        self.master()?.theme()
    }

    /// Get the background defined by this layout.
    ///
    /// Returns `None` when the layout has no local background.
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        if let Some(background) = resolve_picture_background(self.part.part(), self.package)? {
            return Ok(Some(background));
        }

        self.part.background()
    }

    /// Get the background this layout will use after inheritance is applied.
    ///
    /// A background defined on the layout takes precedence over its slide
    /// master. If the layout has no local background, the master's background
    /// is returned instead.
    pub fn effective_background(
        &self,
    ) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        if let Some(background) = self.background()? {
            return Ok(Some(background));
        }

        self.master()?.background()
    }

    /// Get access to the underlying layout part.
    #[inline]
    pub fn part(&self) -> &SlideLayoutPart<'a> {
        &self.part
    }
}

/// A slide master.
///
/// Slide masters define the overall theme and default formatting for slides.
/// Each slide master can have multiple slide layouts.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi_ooxml::pptx::Package;
///
/// let pkg = Package::open("presentation.pptx")?;
/// let pres = pkg.presentation()?;
///
/// for master in pres.slide_masters()?.iter() {
///     println!("Master name: {}", master.name()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct SlideMaster<'a> {
    /// The underlying slide master part
    part: SlideMasterPart<'a>,
    /// Package used to resolve the master-owned layout relationships.
    package: &'a OpcPackage,
}

impl<'a> SlideMaster<'a> {
    /// Create a new SlideMaster with its owning package.
    #[inline]
    pub(crate) fn with_package(part: SlideMasterPart<'a>, package: &'a OpcPackage) -> Self {
        Self { part, package }
    }

    /// Get the master name.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let masters = pres.slide_masters()?;
    ///
    /// if let Some(master) = masters.first() {
    ///     println!("First master name: {}", master.name()?);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn name(&self) -> Result<String> {
        self.part.name()
    }

    /// Get all shapes defined by this slide master.
    pub fn shapes(&self) -> Result<Vec<BaseShape>> {
        self.part.shapes()
    }

    /// Get all placeholder shapes defined by this slide master.
    pub fn placeholders(&self) -> Result<Vec<BaseShape>> {
        self.part.placeholders()
    }

    /// Get the transition effect inherited from this master.
    ///
    /// Returns `None` if the master has no transition.
    pub fn transition(&self) -> Result<Option<crate::pptx::transitions::SlideTransition>> {
        self.part.transition()
    }

    /// Get the background defined by this slide master.
    ///
    /// Returns `None` when the master has no local background.
    pub fn background(&self) -> Result<Option<crate::pptx::backgrounds::SlideBackground>> {
        if let Some(background) = resolve_picture_background(self.part.part(), self.package)? {
            return Ok(Some(background));
        }

        self.part.background()
    }

    /// Resolve the theme used by this slide master.
    ///
    /// The theme relationship must be internal and target an Office theme part.
    /// A master with no theme relationship or more than one theme relationship
    /// is invalid.
    pub fn theme(&self) -> Result<crate::pptx::parts::Theme> {
        let master_part = self.part.part();
        let mut theme_relationships = master_part.rels().iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                rt::THEME | STRICT_THEME_RELATIONSHIP_TYPE
            )
        });
        let relationship = theme_relationships.next().ok_or_else(|| {
            OoxmlError::InvalidRelationship(
                "slide master does not have a theme relationship".to_string(),
            )
        })?;
        if theme_relationships.next().is_some() {
            return Err(OoxmlError::InvalidRelationship(
                "slide master has multiple theme relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "theme relationship '{}' must be internal",
                relationship.r_id()
            )));
        }

        let part_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "invalid theme relationship '{}': {error}",
                relationship.r_id()
            ))
        })?;
        let theme_part = self.package.get_part(&part_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "theme relationship '{}' targets missing part '{}': {error}",
                relationship.r_id(),
                part_name.as_str()
            ))
        })?;
        if theme_part.content_type() != ct::OFC_THEME {
            return Err(OoxmlError::InvalidContentType {
                expected: ct::OFC_THEME.to_string(),
                got: theme_part.content_type().to_string(),
            });
        }

        ThemePart::from_part(theme_part)?.theme()
    }

    /// Get the relationship IDs of all slide layouts in this master.
    ///
    /// Returns a vector of relationship IDs that can be used to access
    /// the actual slide layout parts.
    pub fn slide_layout_rids(&self) -> Result<Vec<String>> {
        self.part.slide_layout_rids()
    }

    /// Resolve all slide layouts owned by this master in relationship order.
    ///
    /// Layout relationships must be internal PresentationML slide-layout parts.
    pub fn slide_layouts(&self) -> Result<Vec<SlideLayout<'a>>> {
        let package = self.package;
        let master_part = self.part.part();
        let relationship_ids = self.part.slide_layout_rids()?;
        let mut layouts = Vec::with_capacity(relationship_ids.len());

        for relationship_id in relationship_ids {
            let relationship = master_part.rels().get(&relationship_id).ok_or_else(|| {
                OoxmlError::InvalidRelationship(format!(
                    "slide master references missing slide-layout relationship '{relationship_id}'"
                ))
            })?;
            if relationship.is_external() {
                return Err(OoxmlError::InvalidRelationship(format!(
                    "slide-layout relationship '{relationship_id}' must be internal"
                )));
            }
            if !matches!(
                relationship.reltype(),
                rt::SLIDE_LAYOUT | STRICT_SLIDE_LAYOUT_RELATIONSHIP_TYPE
            ) {
                return Err(OoxmlError::InvalidRelationship(format!(
                    "relationship '{relationship_id}' is not a slide-layout relationship"
                )));
            }

            let part_name = relationship.target_partname().map_err(|error| {
                OoxmlError::InvalidRelationship(format!(
                    "invalid slide-layout relationship '{relationship_id}': {error}"
                ))
            })?;
            let layout_part = package.get_part(&part_name).map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "slide-layout relationship '{relationship_id}' targets missing part '{}': {error}",
                    part_name.as_str()
                ))
            })?;
            if layout_part.content_type() != ct::PML_SLIDE_LAYOUT {
                return Err(OoxmlError::InvalidContentType {
                    expected: ct::PML_SLIDE_LAYOUT.to_string(),
                    got: layout_part.content_type().to_string(),
                });
            }
            layouts.push(SlideLayout::with_package(
                SlideLayoutPart::from_part(layout_part)?,
                package,
            ));
        }

        Ok(layouts)
    }

    /// Get access to the underlying master part.
    #[inline]
    pub fn part(&self) -> &SlideMasterPart<'a> {
        &self.part
    }
}

#[cfg(test)]
mod tests {
    // Tests will be added as implementation progresses
}
