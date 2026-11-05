//! Presentation slide wrappers: [`Slide`], [`SlideData`] and
//! [`LegacySlideData`].
//!
//! Moved from the facade's `presentation::{slide, types}` (record 0003 §2.4).
//! The [`PresentationView`](crate::PresentationView) family view returns
//! them.
//!
//! **Scheduled for deletion in M1**, with the rest of [`crate::compat`].

use crate::Result;

/// The text and name extracted from a PPTX slide.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct SlideData {
    /// The slide's text, in reading order.
    pub text: String,
    /// The slide's name (`p:cSld@name`).
    pub name: Option<String>,
}

/// The text, one-based number and shape count extracted from a legacy PPT
/// slide.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct LegacySlideData {
    /// The slide's text.
    pub text: String,
    /// The slide's one-based number in the presentation.
    pub slide_number: usize,
    /// The number of shapes on the slide.
    pub shape_count: usize,
}

/// A slide of a presentation.
///
/// **Scheduled for deletion in M1**, with the rest of [`crate::compat`].
#[derive(Debug, Clone)]
#[allow(
    clippy::large_enum_variant,
    reason = "moved unchanged from the facade, where boxing the large variant would have broken the public API"
)]
#[non_exhaustive]
pub enum Slide {
    /// A legacy PowerPoint (`.ppt`) slide.
    #[cfg(feature = "ppt")]
    Ppt(LegacySlideData),
    /// A PresentationML (`.pptx`) slide.
    #[cfg(feature = "pptx")]
    Pptx(SlideData),
    /// An OpenDocument Presentation slide.
    #[cfg(feature = "odp")]
    Odp(litchi_odp::Slide),
}

impl Slide {
    /// The text content of the slide.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the surface uniform.
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "ppt")]
            Self::Ppt(data) => Ok(data.text.clone()),
            #[cfg(feature = "pptx")]
            Self::Pptx(data) => Ok(data.text.clone()),
            #[cfg(feature = "odp")]
            Self::Odp(slide) => Ok(slide.all_text()),
        }
    }

    /// The slide's one-based position in the presentation, for legacy
    /// PowerPoint slides; `None` for PPTX and ODP slides.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn number(&self) -> Option<usize> {
        match self {
            #[cfg(feature = "ppt")]
            Self::Ppt(data) => Some(data.slide_number),
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => None,
            #[cfg(feature = "odp")]
            Self::Odp(_) => None,
        }
    }

    /// The title visible on the slide canvas, for ODP slides with a semantic
    /// title; `None` for legacy PowerPoint and PPTX slides, whose extracted
    /// data does not model a distinct title.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// Returns the ODP typed error when the slide's title cannot be read.
    pub fn title(&self) -> Result<Option<String>> {
        match self {
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => Ok(None),
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => Ok(None),
            #[cfg(feature = "odp")]
            Self::Odp(slide) => Ok(slide
                .title()
                .map_err(|error| {
                    crate::Error::from(error).with_format(litchi_core::FileFormat::Odp)
                })?
                .map(str::to_owned)),
        }
    }

    /// The number of shapes on the slide, for legacy PowerPoint slides;
    /// `None` for PPTX and ODP slides.
    ///
    /// Scheduled for deletion in M1.
    #[must_use]
    pub fn shape_count(&self) -> Option<usize> {
        match self {
            #[cfg(feature = "ppt")]
            Self::Ppt(data) => Some(data.shape_count),
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => None,
            #[cfg(feature = "odp")]
            Self::Odp(_) => None,
        }
    }

    /// The slide's non-content name, for PPTX slides; `None` when the format
    /// does not expose a name or the slide has none.
    ///
    /// Scheduled for deletion in M1.
    ///
    /// # Errors
    ///
    /// None today; the `Result` keeps the surface uniform.
    pub fn name(&self) -> Result<Option<String>> {
        match self {
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => Ok(None),
            #[cfg(feature = "pptx")]
            Self::Pptx(data) => Ok(data.name.clone()),
            #[cfg(feature = "odp")]
            Self::Odp(_) => Ok(None),
        }
    }
}
