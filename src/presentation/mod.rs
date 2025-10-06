/// Unified PowerPoint presentation module.
///
/// This module provides a unified API for working with PowerPoint presentations in both
/// legacy (.ppt) and modern (.pptx) formats. The format is automatically detected
/// and handled transparently.
///
/// # Architecture
///
/// The module provides a format-agnostic API following the python-pptx design:
/// - `Presentation`: The main presentation API (auto-detects format)
/// - `Slide`: Individual slide with shapes and content
/// - `Shape`: Shape elements on slides
///
/// # Example
///
/// ```rust,no_run
/// use litchi::Presentation;
///
/// // Open any PowerPoint presentation (.ppt or .pptx) - format auto-detected
/// let pres = Presentation::open("presentation.ppt")?;
///
/// // Extract all text
/// let text = pres.text()?;
/// println!("Presentation text: {}", text);
///
/// // Get slide count
/// let count = pres.slide_count()?;
/// println!("Slides: {}", count);
///
/// // Access slides
/// for (i, slide) in pres.slides()?.iter().enumerate() {
///     println!("Slide {}: {}", i + 1, slide.text()?);
/// }
/// # Ok::<(), litchi::common::Error>(())
/// ```

use crate::common::{Error, Result};
use crate::ole;
use crate::ooxml;
use std::fs::File;
use std::io::{Cursor, Read, Seek};
use std::path::Path;

/// Extracted data from a PPTX slide (to avoid lifetime issues).
#[derive(Debug, Clone)]
pub struct PptxSlideData {
    text: String,
    name: Option<String>,
}

/// A PowerPoint presentation that can be either .ppt or .pptx format.
///
/// This enum wraps the format-specific implementations and provides
/// a unified API. Users typically don't interact with this enum directly,
/// but instead use the methods on `Presentation`.
enum PresentationImpl {
    /// Legacy .ppt format
    Ppt(Box<ole::ppt::Presentation>),
    /// Modern .pptx format
    Pptx(Box<ooxml::pptx::Presentation<'static>>),
}

/// A PowerPoint presentation.
///
/// This is the main entry point for working with PowerPoint presentations.
/// It automatically detects whether the file is .ppt or .pptx format
/// and provides a unified API.
///
/// Not intended to be constructed directly. Use `Presentation::open()` to
/// open a presentation.
///
/// # Examples
///
/// ```rust,no_run
/// use litchi::Presentation;
///
/// // Open a presentation (format auto-detected)
/// let pres = Presentation::open("slides.ppt")?;
///
/// // Get slide count
/// let count = pres.slide_count()?;
/// println!("Slides: {}", count);
///
/// // Extract text
/// let text = pres.text()?;
/// println!("{}", text);
/// # Ok::<(), litchi::common::Error>(())
/// ```
pub struct Presentation {
    /// The underlying format-specific implementation
    inner: PresentationImpl,
    /// PPTX package storage that must outlive the Presentation reference.
    ///
    /// This field is prefixed with `_` because it's not directly accessed,
    /// but it MUST be kept to maintain memory safety. The `inner` PresentationImpl::Pptx
    /// variant holds a reference with extended lifetime to data owned by this Box.
    /// Dropping this would invalidate those references (use-after-free).
    ///
    /// Only used for PPTX files; None for PPT files.
    _package: Option<Box<ooxml::pptx::Package>>,
}

impl Presentation {
    /// Open a PowerPoint presentation from a file path.
    ///
    /// The file format (.ppt or .pptx) is automatically detected by examining
    /// the file header. You don't need to specify the format explicitly.
    ///
    /// # Arguments
    ///
    /// * `path` - Path to the PowerPoint presentation
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// // Open a .ppt file
    /// let pres1 = Presentation::open("legacy.ppt")?;
    ///
    /// // Open a .pptx file
    /// let pres2 = Presentation::open("modern.pptx")?;
    ///
    /// // Both work the same way
    /// println!("Pres 1: {} slides", pres1.slide_count()?);
    /// println!("Pres 2: {} slides", pres2.slide_count()?);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let path = path.as_ref();
        
        // Try to detect the format by reading the file header
        let mut file = File::open(path)?;
        let format = detect_presentation_format(&mut file)?;
        
        // Open with the appropriate parser
        match format {
            PresentationFormat::Ppt => {
                let mut package = ole::ppt::Package::open(path)
                    .map_err(Error::from)?;
                let pres = package.presentation()
                    .map_err(Error::from)?;
                
                Ok(Self {
                    inner: PresentationImpl::Ppt(Box::new(pres)),
                    _package: None,
                })
            }
            PresentationFormat::Pptx => {
                let package = Box::new(ooxml::pptx::Package::open(path)
                    .map_err(Error::from)?);
                
                // SAFETY: We're using unsafe here to extend the lifetime of the presentation
                // reference. This is safe because we're storing the package in the same
                // struct, ensuring it lives as long as the presentation reference.
                let pres_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::pptx::Package;
                    let pres = (*pkg_ptr).presentation()
                        .map_err(Error::from)?;
                    std::mem::transmute::<ooxml::pptx::Presentation<'_>, ooxml::pptx::Presentation<'static>>(pres)
                };
                
                Ok(Self {
                    inner: PresentationImpl::Pptx(Box::new(pres_ref)),
                    _package: Some(package),
                })
            }
        }
    }

    /// Create a Presentation from a byte buffer.
    ///
    /// This method is optimized for parsing presentations from memory, such as
    /// from network traffic or in-memory caches, without creating temporary files.
    /// It automatically detects the format (.ppt or .pptx) from the byte signature.
    ///
    /// # Arguments
    ///
    /// * `bytes` - The presentation bytes
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    /// use std::fs;
    ///
    /// // From owned bytes (e.g., network data)
    /// let data = fs::read("presentation.ppt")?;
    /// let pres = Presentation::from_bytes(data)?;
    /// println!("Slides: {}", pres.slide_count()?);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    ///
    /// # Performance Notes
    ///
    /// - For .ppt files (OLE2): Parses directly from the buffer with minimal copying
    /// - For .pptx files (ZIP): Efficient decompression without file I/O overhead
    /// - Ideal for network data, streams, or in-memory content
    /// - No temporary files created
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Detect format from byte signature
        let format = detect_presentation_format_from_bytes(&bytes)?;
        
        match format {
            PresentationFormat::Ppt => {
                // For OLE2, create cursor from bytes
                let cursor = Cursor::new(bytes);
                
                let mut package = ole::ppt::Package::from_reader(cursor)
                    .map_err(Error::from)?;
                let pres = package.presentation()
                    .map_err(Error::from)?;
                
                Ok(Self {
                    inner: PresentationImpl::Ppt(Box::new(pres)),
                    _package: None,
                })
            }
            PresentationFormat::Pptx => {
                // For OOXML/ZIP, Cursor<Vec<u8>> implements Read + Seek
                let cursor = Cursor::new(bytes);
                
                let package = Box::new(ooxml::pptx::Package::from_reader(cursor)
                    .map_err(Error::from)?);
                
                // SAFETY: Same lifetime extension as in `open()`
                let pres_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::pptx::Package;
                    let pres = (*pkg_ptr).presentation()
                        .map_err(Error::from)?;
                    std::mem::transmute::<ooxml::pptx::Presentation<'_>, ooxml::pptx::Presentation<'static>>(pres)
                };
                
                Ok(Self {
                    inner: PresentationImpl::Pptx(Box::new(pres_ref)),
                    _package: Some(package),
                })
            }
        }
    }

    /// Get all text content from the presentation.
    ///
    /// This extracts all text from all slides in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// let text = pres.text()?;
    /// println!("{}", text);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn text(&self) -> Result<String> {
        match &self.inner {
            PresentationImpl::Ppt(pres) => {
                pres.text().map_err(Error::from)
            }
            PresentationImpl::Pptx(pres) => {
                // PPTX presentations need to extract text from all slides
                let slides = pres.slides().map_err(Error::from)?;
                let mut texts = Vec::new();
                for slide in slides {
                    if let Ok(text) = slide.text() {
                        if !text.is_empty() {
                            texts.push(text);
                        }
                    }
                }
                Ok(texts.join("\n\n"))
            }
        }
    }

    /// Get the number of slides in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// let count = pres.slide_count()?;
    /// println!("Slides: {}", count);
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn slide_count(&self) -> Result<usize> {
        match &self.inner {
            PresentationImpl::Ppt(pres) => {
                pres.slide_count().map_err(Error::from)
            }
            PresentationImpl::Pptx(pres) => {
                pres.slide_count().map_err(Error::from)
            }
        }
    }

    /// Get the slides in the presentation.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.ppt")?;
    /// for slide in pres.slides()? {
    ///     println!("Slide: {}", slide.text()?);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn slides(&self) -> Result<Vec<Slide>> {
        match &self.inner {
            PresentationImpl::Ppt(pres) => {
                let slides = pres.slides()
                    .map_err(Error::from)?;
                Ok(slides.into_iter().map(Slide::Ppt).collect())
            }
            PresentationImpl::Pptx(pres) => {
                let slides = pres.slides()
                    .map_err(Error::from)?;
                // Extract slide data immediately to avoid lifetime issues
                slides.iter()
                    .map(|s| {
                        let text = s.text().map_err(Error::from)?;
                        let name = s.name().ok();
                        Ok(Slide::Pptx(PptxSlideData { text, name }))
                    })
                    .collect()
            }
        }
    }

    /// Get the slide width in EMUs (English Metric Units).
    ///
    /// Only available for .pptx format. Returns None for .ppt files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.pptx")?;
    /// if let Some(width) = pres.slide_width()? {
    ///     println!("Slide width: {} EMUs", width);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn slide_width(&self) -> Result<Option<i64>> {
        match &self.inner {
            PresentationImpl::Ppt(_) => Ok(None),
            PresentationImpl::Pptx(pres) => {
                pres.slide_width().map_err(Error::from)
            }
        }
    }

    /// Get the slide height in EMUs (English Metric Units).
    ///
    /// Only available for .pptx format. Returns None for .ppt files.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.pptx")?;
    /// if let Some(height) = pres.slide_height()? {
    ///     println!("Slide height: {} EMUs", height);
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn slide_height(&self) -> Result<Option<i64>> {
        match &self.inner {
            PresentationImpl::Ppt(_) => Ok(None),
            PresentationImpl::Pptx(pres) => {
                pres.slide_height().map_err(Error::from)
            }
        }
    }
}

/// A slide in a PowerPoint presentation.
pub enum Slide {
    Ppt(ole::ppt::Slide),
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
            Slide::Ppt(s) => s.text().map_err(Error::from),
            Slide::Pptx(data) => Ok(data.text.clone()),
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

/// Presentation format detection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PresentationFormat {
    /// Legacy .ppt format (OLE2)
    Ppt,
    /// Modern .pptx format (OOXML/ZIP)
    Pptx,
}

/// Detect the presentation format by reading the file header.
///
/// This function reads the first few bytes of the file to determine if it's
/// an OLE2 file (.ppt) or a ZIP file (.pptx).
fn detect_presentation_format<R: Read + Seek>(reader: &mut R) -> Result<PresentationFormat> {
    use std::io::SeekFrom;

    // Read the first 8 bytes
    let mut header = [0u8; 8];
    reader.read_exact(&mut header)?;
    
    // Reset to the beginning
    reader.seek(SeekFrom::Start(0))?;

    detect_presentation_format_from_signature(&header)
}

/// Detect the presentation format from a byte buffer.
///
/// This is optimized for in-memory detection without seeking.
#[inline]
fn detect_presentation_format_from_bytes(bytes: &[u8]) -> Result<PresentationFormat> {
    if bytes.len() < 4 {
        return Err(Error::InvalidFormat("File too small to determine format".to_string()));
    }
    
    detect_presentation_format_from_signature(&bytes[0..8.min(bytes.len())])
}

/// Detect format from the signature bytes.
#[inline]
fn detect_presentation_format_from_signature(header: &[u8]) -> Result<PresentationFormat> {
    // Check for OLE2 signature (D0 CF 11 E0 A1 B1 1A E1)
    if header.len() >= 4 && header[0..4] == [0xD0, 0xCF, 0x11, 0xE0] {
        return Ok(PresentationFormat::Ppt);
    }

    // Check for ZIP signature (PK\x03\x04)
    if header.len() >= 4 && header[0..4] == [0x50, 0x4B, 0x03, 0x04] {
        return Ok(PresentationFormat::Pptx);
    }

    Err(Error::NotOfficeFile)
}

