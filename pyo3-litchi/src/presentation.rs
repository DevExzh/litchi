//! Presentation API bindings

use pyo3::prelude::*;
use pyo3::types::PyModule;
use std::path::PathBuf;
use std::sync::Arc;

use crate::common::to_py_err;

/// Registers presentation types with the Python module
pub fn register(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_class::<Presentation>()?;
    m.add_class::<Slide>()?;
    Ok(())
}

/// Unified PowerPoint presentation interface
///
/// Provides format-agnostic interface for both .ppt and .pptx files.
/// The format is automatically detected when opening a file.
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
#[pyclass(unsendable)]
pub struct Presentation {
    inner: Arc<litchi::Presentation>,
}

#[pymethods]
impl Presentation {
    /// Open a PowerPoint presentation from a file path
    ///
    /// The file format (.ppt or .pptx) is automatically detected.
    ///
    /// Args:
    ///     path: Path to the presentation file
    ///
    /// Returns:
    ///     Presentation instance
    ///
    /// Raises:
    ///     IOError: If the file cannot be read
    ///     ValueError: If the file format is invalid or unsupported
    #[staticmethod]
    fn open(path: PathBuf) -> PyResult<Self> {
        let pres = litchi::Presentation::open(path).map_err(to_py_err)?;
        Ok(Presentation {
            inner: Arc::new(pres),
        })
    }

    /// Extract all text from the presentation
    ///
    /// Returns:
    ///     All text content from all slides as a single string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    /// Get the number of slides in the presentation
    ///
    /// Returns:
    ///     Number of slides
    fn slide_count(&self) -> PyResult<usize> {
        self.inner.slide_count().map_err(to_py_err)
    }

    /// Get all slides in the presentation
    ///
    /// Returns:
    ///     List of Slide objects
    fn slides(&self) -> PyResult<Vec<Slide>> {
        let slides = self.inner.slides().map_err(to_py_err)?;
        Ok(slides
            .into_iter()
            .map(|s| Slide { inner: Arc::new(s) })
            .collect())
    }

    fn __repr__(&self) -> PyResult<String> {
        let slide_count = self.slide_count().unwrap_or(0);
        Ok(format!("<Presentation: {} slides>", slide_count))
    }
}

/// A slide in a presentation
///
/// Represents a single slide with text and shapes.
#[pyclass(unsendable)]
pub struct Slide {
    inner: Arc<litchi::presentation::Slide>,
}

#[pymethods]
impl Slide {
    /// Extract all text from the slide
    ///
    /// Returns:
    ///     All text content as a single string
    fn text(&self) -> PyResult<String> {
        self.inner.text().map_err(to_py_err)
    }

    fn __repr__(&self) -> PyResult<String> {
        let text = self.text().unwrap_or_default();
        let preview = if text.len() > 50 {
            format!("{}...", &text[..50])
        } else {
            text
        };
        Ok(format!("<Slide: '{}'>", preview))
    }
}
