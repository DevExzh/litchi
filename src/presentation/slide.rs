//! Slide implementation for PowerPoint presentations.

use crate::common::Result;
use super::types::{PptxSlideData, PptSlideData};

/// A slide in a PowerPoint presentation.
pub enum Slide {
    /// Legacy PPT slide with extracted data
    Ppt(PptSlideData),
    /// Modern PPTX slide with extracted data
    Pptx(PptxSlideData),
}

impl Slide {
    /// Get the text content of the slide.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// for slide in pres.slides()? {
    ///     println!("Slide text: {}", slide.text()?);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        match self {
            Slide::Ppt(data) => Ok(data.text.clone()),
            Slide::Pptx(data) => Ok(data.text.clone()),
        }
    }

    /// Get the slide number (1-based).
    ///
    /// Only available for .ppt format. Returns None for .pptx files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// for slide in pres.slides()? {
    ///     if let Some(num) = slide.number() {
    ///         println!("Slide number: {}", num);
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn number(&self) -> Option<usize> {
        match self {
            Slide::Ppt(data) => Some(data.slide_number),
            Slide::Pptx(_) => None,
        }
    }

    /// Get the number of shapes on the slide.
    ///
    /// Only available for .ppt format. Returns None for .pptx files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// for slide in pres.slides()? {
    ///     if let Some(count) = slide.shape_count() {
    ///         println!("Shapes: {}", count);
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn shape_count(&self) -> Option<usize> {
        match self {
            Slide::Ppt(data) => Some(data.shape_count),
            Slide::Pptx(_) => None,
        }
    }

    /// Get the slide name.
    ///
    /// Only available for .pptx format. Returns None for .ppt files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.pptx")?;
    /// for slide in pres.slides()? {
    ///     if let Some(name) = slide.name()? {
    ///         println!("Slide name: {}", name);
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn name(&self) -> Result<Option<String>> {
        match self {
            Slide::Ppt(_) => Ok(None),
            Slide::Pptx(data) => Ok(data.name.clone()),
        }
    }
}

