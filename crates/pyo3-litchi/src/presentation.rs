//! Presentation API bindings

use std::path::PathBuf;

use litchi::FileFormat as CoreFormat;
use litchi::compat::slide;
use pyo3::prelude::*;
use pyo3::types::PyModule;

use crate::common::preview;
use crate::errors::{Failure, Opener, to_py_err};
use crate::iwork;

/// Registers presentation types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Presentation>()?;
    m.add_class::<Slide>()?;
    Ok(())
}

/// What a `Presentation` reads: a unified file of the presentation family,
/// or a Keynote `File`, which has no family view (its leaf package is read).
enum Source {
    File(litchi::File),
    Keynote(litchi::File),
}

/// Unified PowerPoint presentation interface
///
/// Provides a format-agnostic interface for presentations: .ppt, .pptx,
/// .odp and .key. The format is detected from the file's content when it is
/// opened.
///
/// A Presentation is immutable and can be shared between threads. Opening
/// and the presentation-wide reads release the GIL while they run.
///
/// # Examples
///
/// ```python
/// from litchi_py import Presentation
///
/// # Open any PowerPoint presentation (.ppt or .pptx)
/// pres = Presentation.open("presentation.pptx")
///
/// # Extract all text
/// text = pres.text()
/// print(text)
///
/// # Get slide count
/// print(f"Total slides: {pres.slide_count()}")
///
/// # Access individual slides
/// for i, slide in enumerate(pres.slides()):
///     print(f"Slide {i + 1}: {slide.text()}")
/// ```
#[pyclass(frozen, module = "litchi_py")]
pub struct Presentation {
    source: Source,
}

impl Presentation {
    fn read(path: &std::path::Path) -> Result<Source, Failure> {
        match litchi::open(path) {
            // `litchi::File` opens Keynote packages and package directories
            // (records 0296 and 0306), but its family views do not cover
            // them; this reader projects the package the file retains.
            Ok(file) if file.format() == Some(CoreFormat::Keynote) => Ok(Source::Keynote(file)),
            Ok(file) => match file.presentation() {
                Ok(_) => Ok(Source::File(file)),
                Err(error) => Err(Failure::open(&error, Opener::Presentation)),
            },
            Err(error) => Err(Failure::open(&error, Opener::Presentation)),
        }
    }

    fn count(&self) -> Result<usize, Failure> {
        match &self.source {
            Source::File(file) => file
                .presentation()
                .and_then(|presentation| presentation.slide_count())
                .map_err(|error| Failure::read(&error)),
            Source::Keynote(file) => iwork::keynote(file)?
                .slides()
                .map(<[_]>::len)
                .map_err(|error| iwork::keynote_failure(&error)),
        }
    }
}

#[pymethods]
impl Presentation {
    /// Open a PowerPoint presentation from a file path
    ///
    /// The format (.ppt, .pptx, .odp or .key) is detected from the file's
    /// content, not from its name.
    ///
    /// Args:
    ///     path: Path to the presentation file
    ///
    /// Returns:
    ///     Presentation instance
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file cannot be read
    ///     UnsupportedError: (a ValueError) If the file is not a
    ///         presentation, is not a recognized format, or uses a variant
    ///         that cannot be read (flat OpenDocument)
    ///     CorruptFileError: (a ValueError) If the file is malformed
    ///     ProtectedError: (a ValueError) If the presentation is encrypted
    #[staticmethod]
    fn open(py: Python<'_>, path: PathBuf) -> PyResult<Self> {
        let source = py.detach(|| Self::read(&path))?;
        Ok(Presentation { source })
    }

    /// Extract all text from the presentation
    ///
    /// Returns:
    ///     All text content from all slides as a single string
    ///
    /// Raises:
    ///     IoError: (an OSError) If the file changed after it was opened
    fn text(&self, py: Python<'_>) -> PyResult<String> {
        py.detach(|| match &self.source {
            Source::File(file) => file.text().map_err(|error| Failure::read(&error)),
            Source::Keynote(file) => file.text().map_err(|error| Failure::read(&error)),
        })
        .map_err(PyErr::from)
    }

    /// Get the number of slides in the presentation
    ///
    /// Returns:
    ///     Number of slides
    fn slide_count(&self, py: Python<'_>) -> PyResult<usize> {
        py.detach(|| self.count()).map_err(PyErr::from)
    }

    /// Get all slides in the presentation
    ///
    /// Returns:
    ///     List of Slide objects
    fn slides(&self, py: Python<'_>) -> PyResult<Vec<Slide>> {
        let slides = py.detach(|| match &self.source {
            Source::File(file) => file
                .presentation()
                .and_then(|presentation| presentation.slides())
                .map(|slides| slides.into_iter().map(SlideInner::Compat).collect())
                .map_err(|error| Failure::read(&error)),
            Source::Keynote(file) => iwork::keynote(file)?
                .slides()
                .map(|slides| {
                    slides
                        .iter()
                        .map(|slide| SlideInner::Keynote(slide.plain_text()))
                        .collect::<Vec<_>>()
                })
                .map_err(|error| iwork::keynote_failure(&error)),
        })?;
        Ok(slides.into_iter().map(|inner| Slide { inner }).collect())
    }

    fn __repr__(&self, py: Python<'_>) -> PyResult<String> {
        let slide_count = py.detach(|| self.count()).unwrap_or(0);
        Ok(format!("<Presentation: {} slides>", slide_count))
    }
}

#[allow(
    clippy::large_enum_variant,
    reason = "each value is the payload of one Python object, and all but the iWork case hold the large variant, so boxing it would add an allocation per element and save nothing"
)]
enum SlideInner {
    Compat(slide::Slide),
    /// The text of a Keynote slide.
    Keynote(String),
}

/// A slide in a presentation
///
/// Represents a single slide with text and shapes.
#[pyclass(frozen, module = "litchi_py")]
pub struct Slide {
    inner: SlideInner,
}

#[pymethods]
impl Slide {
    /// Extract all text from the slide
    ///
    /// Returns:
    ///     All text content as a single string
    fn text(&self) -> PyResult<String> {
        match &self.inner {
            SlideInner::Compat(slide) => slide.text().map_err(to_py_err),
            SlideInner::Keynote(text) => Ok(text.clone()),
        }
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        Ok(format!("<Slide: '{}'>", preview(&text, 50)))
    }
}
