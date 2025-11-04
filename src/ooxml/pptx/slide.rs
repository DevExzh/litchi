/// Slide-related objects, including Slide, SlideLayout, and SlideMaster.
use crate::ooxml::error::Result;
use crate::ooxml::opc::packuri::PackURI;
use crate::ooxml::pptx::parts::{SlideLayoutPart, SlideMasterPart, SlidePart};
use crate::ooxml::pptx::shapes::base::BaseShape;

/// A slide in a presentation.
///
/// Provides access to slide content and properties, following the python-pptx
/// interface design.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::ooxml::pptx::Package;
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
    package: Option<&'a crate::ooxml::opc::OpcPackage>,
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
    pub(crate) fn with_package(
        part: SlidePart<'a>,
        package: &'a crate::ooxml::opc::OpcPackage,
    ) -> Self {
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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

    /// Get all shapes on this slide.
    ///
    /// Returns a vector of BaseShape objects that provide access to text,
    /// pictures, tables, and other shape types.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    /// use litchi::ooxml::pptx::shapes::ShapeType;
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

    /// Get the number of shapes on this slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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

    /// Check if the slide has any tables.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
        use crate::ooxml::pptx::shapes::ShapeType;

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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    /// use litchi::ooxml::pptx::Package;
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
    pub fn transition(&self) -> Result<Option<crate::ooxml::pptx::transitions::SlideTransition>> {
        self.part.transition()
    }

    /// Get the background for this slide.
    ///
    /// Returns `None` if no background is defined (uses master background).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
    ///
    /// let pkg = Package::open("presentation.pptx")?;
    /// let pres = pkg.presentation()?;
    /// let slides = pres.slides()?;
    ///
    /// if let Some(slide) = slides.first() {
    ///     if let Some(bg) = slide.background()? {
    ///         println!("Slide has custom background: {:?}", bg);
    ///     } else {
    ///         println!("Using master background");
    ///     }
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn background(&self) -> Result<Option<crate::ooxml::pptx::backgrounds::SlideBackground>> {
        self.part.background()
    }

    /// Get the speaker notes for this slide.
    ///
    /// Returns `None` if no notes are defined or if the package reference is not available.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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
        // Check if we have package reference
        let package = match self.package {
            Some(pkg) => pkg,
            None => return Ok(None),
        };

        // Look for notes relationship
        let slide_part = self.part.part();
        let rels = slide_part.rels();

        // Find the notes relationship (type is notesSlide)
        let notes_rel = rels.iter().find(|rel| {
            rel.reltype()
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
        });

        if let Some(rel) = notes_rel {
            // Get the notes part
            let base_uri = slide_part.partname().base_uri();
            let notes_partname = PackURI::from_rel_ref(base_uri, rel.target_ref())
                .map_err(crate::ooxml::error::OoxmlError::InvalidFormat)?;

            if let Ok(notes_part) = package.get_part(&notes_partname) {
                // Extract text from notes
                return Self::extract_notes_text(notes_part.blob());
            }
        }

        Ok(None)
    }

    /// Extract text from notes XML.
    fn extract_notes_text(xml: &[u8]) -> Result<Option<String>> {
        use quick_xml::Reader;
        use quick_xml::events::Event;

        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut text = String::new();
        let mut in_text_element = false;
        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = true;
                    }
                },
                Ok(Event::Text(e)) if in_text_element => {
                    let t = std::str::from_utf8(e.as_ref())
                        .map_err(|e| crate::ooxml::error::OoxmlError::Xml(e.to_string()))?;
                    if !text.is_empty() && !text.ends_with('\n') {
                        text.push('\n');
                    }
                    text.push_str(t);
                },
                Ok(Event::End(e)) => {
                    if e.local_name().as_ref() == b"t" {
                        in_text_element = false;
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(crate::ooxml::error::OoxmlError::Xml(e.to_string())),
                _ => {},
            }
            buf.clear();
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
}

impl<'a> SlideLayout<'a> {
    /// Create a new SlideLayout from a SlideLayoutPart.
    #[allow(unused)]
    #[inline]
    pub(crate) fn new(part: SlideLayoutPart<'a>) -> Self {
        Self { part }
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
/// use litchi::ooxml::pptx::Package;
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
}

impl<'a> SlideMaster<'a> {
    /// Create a new SlideMaster from a SlideMasterPart.
    #[inline]
    pub(crate) fn new(part: SlideMasterPart<'a>) -> Self {
        Self { part }
    }

    /// Get the master name.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::pptx::Package;
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

    /// Get the relationship IDs of all slide layouts in this master.
    ///
    /// Returns a vector of relationship IDs that can be used to access
    /// the actual slide layout parts.
    pub fn slide_layout_rids(&self) -> Result<Vec<String>> {
        self.part.slide_layout_rids()
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
