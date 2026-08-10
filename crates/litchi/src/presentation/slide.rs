//! Slide implementation for PowerPoint presentations.

use super::types::{LegacySlideData, SlideData};
use litchi_core::Result;

/// A slide in a PowerPoint presentation.
#[allow(
    clippy::large_enum_variant,
    reason = "public facade enum; boxing the large variant would break the API"
)]
pub enum Slide {
    /// Legacy PPT slide with extracted data
    Ppt(LegacySlideData),
    /// Modern PPTX slide with extracted data
    Pptx(SlideData),
    /// Apple Keynote slide
    #[cfg(feature = "keynote")]
    Keynote {
        /// One-based position in the presentation, not a native Keynote object identifier.
        number: usize,
        /// Optional developer-facing name shown in Keynote's slide navigator.
        name: Option<String>,
        /// Optional title that is visible on the slide canvas.
        title: Option<String>,
        /// All modeled slide text in semantic order.
        text: String,
    },
    /// OpenDocument Presentation slide
    #[cfg(feature = "odp")]
    Odp(litchi_odp::Slide),
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
            #[cfg(feature = "keynote")]
            Slide::Keynote { text, .. } => Ok(text.clone()),
            #[cfg(feature = "odp")]
            Slide::Odp(slide) => Ok(slide.all_text()),
        }
    }

    /// Get the slide's one-based position in the presentation.
    ///
    /// Available for legacy PowerPoint and Keynote slides. For Keynote this is
    /// derived from presentation order and is not a native object identifier.
    /// Returns `None` for PPTX and ODP slides.
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
            #[cfg(feature = "keynote")]
            Slide::Keynote { number, .. } => Some(*number),
            #[cfg(feature = "odp")]
            Slide::Odp(_) => None, // Slide numbers not currently exposed for ODP
        }
    }

    /// Get the title visible on the slide canvas.
    ///
    /// Available for Keynote and ODP slides when a semantic title is present.
    /// A Keynote title is distinct from its developer-facing navigator
    /// [`name`](Self::name) and its one-based [`number`](Self::number).
    /// Returns `None` for legacy PowerPoint and PPTX slides because their
    /// current facade data does not model a distinct semantic title.
    ///
    /// # Errors
    ///
    /// Returns an error if the underlying format cannot provide its title.
    pub fn title(&self) -> Result<Option<String>> {
        match self {
            Slide::Ppt(_) | Slide::Pptx(_) => Ok(None),
            #[cfg(feature = "keynote")]
            Slide::Keynote { title, .. } => Ok(title.clone()),
            #[cfg(feature = "odp")]
            Slide::Odp(slide) => Ok(slide.title()?.map(str::to_owned)),
        }
    }

    /// Get the number of shapes on the slide.
    ///
    /// Only available for .ppt format. Returns None for .pptx and .key files.
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
            #[cfg(feature = "keynote")]
            Slide::Keynote { .. } => None, // Shape count not currently exposed for Keynote
            #[cfg(feature = "odp")]
            Slide::Odp(_) => None, // Shape count not currently exposed for ODP
        }
    }

    /// Get the slide's non-content name.
    ///
    /// For Keynote this is the optional developer-facing navigator name, which
    /// is distinct from both the one-based slide number and the title visible
    /// on the slide canvas. Returns `None` when the format does not expose a
    /// name or the slide has no name.
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
            #[cfg(feature = "keynote")]
            Slide::Keynote { name, .. } => Ok(name.clone()),
            #[cfg(feature = "odp")]
            Slide::Odp(_slide) => Ok(None), // ODP slides don't have names in the current API
        }
    }
}

#[cfg(all(test, feature = "keynote"))]
mod keynote_tests {
    use super::Slide;

    #[test]
    fn keynote_identity_fields_remain_semantically_distinct() {
        let slide = Slide::Keynote {
            number: 2,
            name: Some("Agenda".to_owned()),
            title: Some("Quarterly results".to_owned()),
            text: "Quarterly results\nRevenue increased".to_owned(),
        };

        assert_eq!(slide.number(), Some(2));
        assert_eq!(slide.name().unwrap().as_deref(), Some("Agenda"));
        assert_eq!(slide.title().unwrap().as_deref(), Some("Quarterly results"));
        assert_eq!(
            slide.text().unwrap(),
            "Quarterly results\nRevenue increased"
        );
        let Slide::Keynote { title, .. } = slide else {
            unreachable!("test value is a Keynote slide")
        };
        assert_eq!(title.as_deref(), Some("Quarterly results"));
    }
}

#[cfg(all(test, feature = "pptx", feature = "ppt"))]
mod tests {
    use super::super::Presentation;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_slide_text_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides = pres.slides().expect("Failed to get slides");

        for slide in slides {
            let _text = slide.text().expect("Failed to get slide text");
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_slide_text_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let slides = pres.slides().expect("Failed to get slides");

        for slide in slides {
            let _text = slide.text().expect("Failed to get slide text");
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_slide_number_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides = pres.slides().expect("Failed to get slides");

        for slide in &slides {
            let num = slide.number();
            assert!(num.is_some(), "PPT slides should have numbers");
            assert!(num.unwrap() > 0, "Slide number should be positive");
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_slide_shape_count_ppt() {
        // Use SampleShow.ppt to avoid metadata overflow issues
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides = pres.slides().expect("Failed to get slides");

        for slide in &slides {
            // shape_count is only available for PPT format
            let _ = slide.shape_count();
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_slide_name_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let slides = pres.slides().expect("Failed to get slides");

        for slide in slides {
            let _name = slide.name().expect("Failed to get slide name");
        }
    }
}
