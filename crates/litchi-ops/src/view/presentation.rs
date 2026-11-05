//! The presentation family view, ported from the facade's
//! `presentation::Presentation` reads.

use litchi_core::FileFormat;

use crate::Result;
#[cfg(feature = "ppt")]
use crate::compat::slide::LegacySlideData;
use crate::compat::slide::Slide;
#[cfg(feature = "pptx")]
use crate::compat::slide::SlideData;
#[cfg(feature = "odp")]
use crate::format::odp::Odp;
#[cfg(feature = "ppt")]
use crate::format::ppt::Ppt;
#[cfg(feature = "pptx")]
use crate::format::pptx::Pptx;

/// The presentation view of a [`File`](crate::File): the slides of PPT, PPTX
/// and ODP presentations.
///
/// [`File::presentation`](crate::File::presentation) returns it. A
/// source-backed PPTX reads only the slides it is asked for and reports a
/// source changed after open as [`Code::IoError`](crate::Code::IoError).
/// PPTX slide catalogs and slide sizes are format-specific; read them on the
/// `Pptx` handle.
///
/// ```no_run
/// # fn main() -> litchi_ops::Result<()> {
/// let file = litchi_ops::open("deck.pptx")?;
/// for slide in file.presentation()?.slides()? {
///     println!("{}", slide.text()?);
/// }
/// # Ok(())
/// # }
/// ```
#[derive(Clone, Copy)]
pub struct PresentationView<'a>(Inner<'a>);

/// The handle a [`PresentationView`] reads, for the crate's adapters built
/// on the view.
#[derive(Clone, Copy)]
pub(crate) enum Inner<'a> {
    #[cfg(feature = "pptx")]
    Pptx(&'a Pptx),
    #[cfg(feature = "ppt")]
    Ppt(&'a Ppt),
    #[cfg(feature = "odp")]
    Odp(&'a Odp),
}

#[cfg(feature = "pptx")]
fn pptx_slide(slide: crate::format::pptx::Slide) -> Slide {
    let (text, name) = slide.into_parts();
    Slide::Pptx(SlideData {
        text,
        name: Some(name),
    })
}

#[cfg(feature = "ppt")]
fn ppt_slide(slide: &litchi_ppt::Slide<'_>) -> Result<Slide> {
    Ok(Slide::Ppt(LegacySlideData {
        text: slide.text()?.to_string(),
        slide_number: slide.slide_number(),
        shape_count: slide.shape_count().unwrap_or(0),
    }))
}

impl<'a> PresentationView<'a> {
    #[cfg(feature = "pptx")]
    pub(crate) fn pptx(handle: &'a Pptx) -> Self {
        Self(Inner::Pptx(handle))
    }

    #[cfg(feature = "ppt")]
    pub(crate) fn ppt(handle: &'a Ppt) -> Self {
        Self(Inner::Ppt(handle))
    }

    #[cfg(feature = "odp")]
    pub(crate) fn odp(handle: &'a Odp) -> Self {
        Self(Inner::Odp(handle))
    }

    /// The handle this view reads.
    #[cfg_attr(
        not(all(feature = "markdown", feature = "yaml")),
        expect(dead_code, reason = "the Markdown front matter is the only reader")
    )]
    pub(crate) fn inner(&self) -> Inner<'a> {
        self.0
    }

    /// The format of the presentation.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        match self.0 {
            #[cfg(feature = "pptx")]
            Inner::Pptx(_) => FileFormat::Pptx,
            #[cfg(feature = "ppt")]
            Inner::Ppt(_) => FileFormat::Ppt,
            #[cfg(feature = "odp")]
            Inner::Odp(_) => FileFormat::Odp,
        }
    }

    /// The number of slides.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`](crate::Code::IoError) when a source-backed
    /// presentation changed after open, and the format's typed error when its
    /// slide list cannot be read.
    pub fn slide_count(&self) -> Result<usize> {
        match self.0 {
            #[cfg(feature = "pptx")]
            Inner::Pptx(pptx) => pptx.slide_count(),
            #[cfg(feature = "ppt")]
            Inner::Ppt(ppt) => ppt.slide_count(),
            #[cfg(feature = "odp")]
            Inner::Odp(odp) => odp.slide_count(),
        }
    }

    /// Every slide, in presentation order.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_count`], and the format's typed error when a slide
    /// cannot be read.
    pub fn slides(&self) -> Result<Vec<Slide>> {
        match self.0 {
            #[cfg(feature = "pptx")]
            Inner::Pptx(pptx) => Ok(pptx.slides()?.into_iter().map(pptx_slide).collect()),
            #[cfg(feature = "ppt")]
            Inner::Ppt(ppt) => ppt.reader().slides()?.iter().map(ppt_slide).collect(),
            #[cfg(feature = "odp")]
            Inner::Odp(odp) => Ok(odp.slides()?.into_iter().map(Slide::Odp).collect()),
        }
    }

    /// The slide at zero-based `position`, or `None` past the last slide.
    /// A source-backed PPTX reads only the selected slide.
    ///
    /// # Errors
    ///
    /// As [`Self::slides`], for the selected slide.
    pub fn slide(&self, position: impl Into<litchi_core::Position>) -> Result<Option<Slide>> {
        let position = position.into().get();
        match self.0 {
            #[cfg(feature = "pptx")]
            Inner::Pptx(pptx) => Ok(pptx.slide(position)?.map(pptx_slide)),
            #[cfg(feature = "ppt")]
            Inner::Ppt(ppt) => ppt
                .reader()
                .slide_at(position)?
                .map(|slide| ppt_slide(&slide))
                .transpose(),
            #[cfg(feature = "odp")]
            Inner::Odp(odp) => Ok(odp.slide(position)?.map(Slide::Odp)),
        }
    }

    /// Slide text for the Markdown adapter: `(slide number, text)` per slide.
    /// Legacy PowerPoint takes its fast path, which skips shape parsing; the
    /// other formats project [`Self::slides`]. The adapter moves onto the node
    /// view in M1.
    ///
    /// # Errors
    ///
    /// As [`Self::slides`].
    #[cfg(any(
        feature = "markdown",
        all(test, any(feature = "ppt", feature = "pptx"))
    ))]
    pub(crate) fn extract_text_for_markdown(&self) -> Result<Vec<(usize, String)>> {
        match self.0 {
            #[cfg(feature = "ppt")]
            Inner::Ppt(ppt) => Ok(ppt.reader().extract_text_fast()?),
            #[allow(
                unreachable_patterns,
                reason = "a build whose only presentation format is PPT always takes the fast path"
            )]
            _ => self
                .slides()?
                .iter()
                .enumerate()
                .map(|(index, slide)| slide.text().map(|text| (index + 1, text)))
                .collect(),
        }
    }
}

impl std::fmt::Debug for PresentationView<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PresentationView")
            .field("format", &self.format())
            .finish_non_exhaustive()
    }
}
