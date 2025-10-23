//! PowerPoint presentation implementation.

use super::Slide;
use super::types::{
    PptSlideData, PptxSlideData, PresentationFormat, PresentationImpl, detect_presentation_format,
    detect_presentation_format_from_bytes,
};
use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

use std::fs::File;
use std::io::Cursor;
use std::path::Path;

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
    pub(super) inner: PresentationImpl,
    /// PPTX package storage that must outlive the Presentation reference.
    ///
    /// This field is prefixed with `_` because it's not directly accessed,
    /// but it MUST be kept to maintain memory safety. The `inner` PresentationImpl::Pptx
    /// variant holds a reference with extended lifetime to data owned by this Box.
    /// Dropping this would invalidate those references (use-after-free).
    ///
    /// Only used for PPTX files; None for PPT and Keynote files.
    #[cfg(feature = "ooxml")]
    pub(super) _pptx_package: Option<Box<ooxml::pptx::Package>>,
    /// Cached metadata extracted during presentation creation.
    ///
    /// Metadata is extracted once during `open()` or `from_bytes()` and cached here
    /// for efficient access. This avoids needing mutable access during `metadata()` calls.
    pub(super) cached_metadata: Option<crate::common::Metadata>,
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
        let initial_format = detect_presentation_format(&mut file)?;

        // Refine ZIP-based format detection (distinguish PPTX from Keynote)
        let format = super::types::refine_presentation_format(&mut file, initial_format)?;

        // Open with the appropriate parser
        match format {
            #[cfg(feature = "ole")]
            PresentationFormat::Ppt => {
                let mut package = Box::new(ole::ppt::Package::open(path).map_err(Error::from)?);

                // Extract metadata from OLE property streams
                let cached_metadata =
                    package
                        .ole_file()
                        .get_metadata()
                        .ok()
                        .and_then(|ole_metadata| {
                            let metadata: crate::common::Metadata = ole_metadata.into();
                            if metadata.has_data() {
                                Some(metadata)
                            } else {
                                None
                            }
                        });

                let pres = package.presentation().map_err(Error::from)?;

                Ok(Self {
                    inner: PresentationImpl::Ppt(pres),
                    #[cfg(feature = "ooxml")]
                    _pptx_package: None,
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "ole"))]
            PresentationFormat::Ppt => Err(Error::FeatureDisabled("ole".to_string())),
            #[cfg(feature = "ooxml")]
            PresentationFormat::Pptx => {
                let package = Box::new(ooxml::pptx::Package::open(path).map_err(Error::from)?);

                // Extract metadata from OOXML package before transferring ownership
                let cached_metadata =
                    crate::ooxml::metadata::extract_metadata(package.opc_package())
                        .ok()
                        .and_then(|metadata| {
                            if metadata.has_data() {
                                Some(metadata)
                            } else {
                                None
                            }
                        });

                // SAFETY: We're using unsafe here to extend the lifetime of the presentation
                // reference. This is safe because we're storing the package in the same
                // struct, ensuring it lives as long as the presentation reference.
                let pres_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::pptx::Package;
                    let pres = (*pkg_ptr).presentation().map_err(Error::from)?;
                    std::mem::transmute::<
                        ooxml::pptx::Presentation<'_>,
                        ooxml::pptx::Presentation<'static>,
                    >(pres)
                };

                Ok(Self {
                    inner: PresentationImpl::Pptx(Box::new(pres_ref)),
                    _pptx_package: Some(package),
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "ooxml"))]
            PresentationFormat::Pptx => Err(Error::FeatureDisabled("ooxml".to_string())),
            #[cfg(feature = "iwa")]
            PresentationFormat::Keynote => {
                let doc = crate::iwa::keynote::KeynoteDocument::open(path).map_err(|e| {
                    Error::ParseError(format!("Failed to open Keynote presentation: {}", e))
                })?;

                // Extract Keynote metadata from bundle properties
                let cached_metadata = doc.metadata().ok().flatten().and_then(|metadata| {
                    if metadata.has_data() {
                        Some(metadata)
                    } else {
                        None
                    }
                });

                Ok(Self {
                    inner: PresentationImpl::Keynote(doc),
                    #[cfg(feature = "ooxml")]
                    _pptx_package: None,
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "iwa"))]
            PresentationFormat::Keynote => Err(Error::FeatureDisabled("iwa".to_string())),
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
        let initial_format = detect_presentation_format_from_bytes(&bytes)?;

        // Refine ZIP-based format detection for bytes
        let format = if initial_format == PresentationFormat::Pptx {
            // Check if it's a Keynote presentation
            #[cfg(feature = "iwa")]
            {
                let mut cursor = Cursor::new(&bytes);
                super::types::refine_presentation_format(&mut cursor, initial_format)?
            }
            #[cfg(not(feature = "iwa"))]
            initial_format
        } else {
            initial_format
        };

        match format {
            #[cfg(feature = "ole")]
            PresentationFormat::Ppt => {
                // For OLE2, create cursor from bytes
                let cursor = Cursor::new(bytes);

                let mut package =
                    Box::new(ole::ppt::Package::from_reader(cursor).map_err(Error::from)?);

                // Extract metadata from OLE property streams
                let cached_metadata =
                    package
                        .ole_file()
                        .get_metadata()
                        .ok()
                        .and_then(|ole_metadata| {
                            let metadata: crate::common::Metadata = ole_metadata.into();
                            if metadata.has_data() {
                                Some(metadata)
                            } else {
                                None
                            }
                        });

                let pres = package.presentation().map_err(Error::from)?;

                Ok(Self {
                    inner: PresentationImpl::Ppt(pres),
                    #[cfg(feature = "ooxml")]
                    _pptx_package: None,
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "ole"))]
            PresentationFormat::Ppt => Err(Error::FeatureDisabled("ole".to_string())),
            #[cfg(feature = "ooxml")]
            PresentationFormat::Pptx => {
                // For OOXML/ZIP, Cursor<Vec<u8>> implements Read + Seek
                let cursor = Cursor::new(bytes);

                let package =
                    Box::new(ooxml::pptx::Package::from_reader(cursor).map_err(Error::from)?);

                // Extract metadata from OOXML package before transferring ownership
                let cached_metadata =
                    crate::ooxml::metadata::extract_metadata(package.opc_package())
                        .ok()
                        .and_then(|metadata| {
                            if metadata.has_data() {
                                Some(metadata)
                            } else {
                                None
                            }
                        });

                // SAFETY: Same lifetime extension as in `open()`
                let pres_ref = unsafe {
                    let pkg_ptr = &*package as *const ooxml::pptx::Package;
                    let pres = (*pkg_ptr).presentation().map_err(Error::from)?;
                    std::mem::transmute::<
                        ooxml::pptx::Presentation<'_>,
                        ooxml::pptx::Presentation<'static>,
                    >(pres)
                };

                Ok(Self {
                    inner: PresentationImpl::Pptx(Box::new(pres_ref)),
                    _pptx_package: Some(package),
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "ooxml"))]
            PresentationFormat::Pptx => Err(Error::FeatureDisabled("ooxml".to_string())),
            #[cfg(feature = "iwa")]
            PresentationFormat::Keynote => {
                let doc =
                    crate::iwa::keynote::KeynoteDocument::from_bytes(&bytes).map_err(|e| {
                        Error::ParseError(format!("Failed to open Keynote from bytes: {}", e))
                    })?;

                // Extract Keynote metadata from bundle properties
                let cached_metadata = doc.metadata().ok().flatten().and_then(|metadata| {
                    if metadata.has_data() {
                        Some(metadata)
                    } else {
                        None
                    }
                });

                Ok(Self {
                    inner: PresentationImpl::Keynote(doc),
                    #[cfg(feature = "ooxml")]
                    _pptx_package: None,
                    cached_metadata,
                })
            },
            #[cfg(not(feature = "iwa"))]
            PresentationFormat::Keynote => Err(Error::FeatureDisabled("iwa".to_string())),
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
            #[cfg(feature = "ole")]
            PresentationImpl::Ppt(pres) => pres.text().map_err(Error::from),
            #[cfg(feature = "ooxml")]
            PresentationImpl::Pptx(pres) => {
                // PPTX presentations need to extract text from all slides
                let slides = pres.slides().map_err(Error::from)?;
                let mut texts = Vec::new();
                for slide in slides {
                    if let Ok(text) = slide.text()
                        && !text.is_empty()
                    {
                        texts.push(text);
                    }
                }
                Ok(texts.join("\n\n"))
            },
            #[cfg(feature = "iwa")]
            PresentationImpl::Keynote(doc) => doc.text().map_err(|e| {
                Error::ParseError(format!("Failed to extract text from Keynote: {}", e))
            }),
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
            #[cfg(feature = "ole")]
            PresentationImpl::Ppt(pres) => Ok(pres.slide_count()),
            #[cfg(feature = "ooxml")]
            PresentationImpl::Pptx(pres) => pres.slide_count().map_err(Error::from),
            #[cfg(feature = "iwa")]
            PresentationImpl::Keynote(doc) => {
                let slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get slides: {}", e)))?;
                Ok(slides.len())
            },
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
            #[cfg(feature = "ole")]
            PresentationImpl::Ppt(pres) => {
                // Extract slide data to avoid lifetime issues
                let ppt_slides = pres.slides().map_err(Error::from)?;
                ppt_slides
                    .iter()
                    .map(|s| {
                        let text = s.text().map_err(Error::from)?.to_string();
                        let slide_number = s.slide_number();
                        let shape_count = s.shape_count().unwrap_or(0);
                        Ok(Slide::Ppt(PptSlideData {
                            text,
                            slide_number,
                            shape_count,
                        }))
                    })
                    .collect()
            },
            #[cfg(feature = "ooxml")]
            PresentationImpl::Pptx(pres) => {
                let slides = pres.slides().map_err(Error::from)?;
                // Extract slide data immediately to avoid lifetime issues
                slides
                    .iter()
                    .map(|s| {
                        let text = s.text().map_err(Error::from)?;
                        let name = s.name().ok();
                        Ok(Slide::Pptx(PptxSlideData { text, name }))
                    })
                    .collect()
            },
            #[cfg(feature = "iwa")]
            PresentationImpl::Keynote(doc) => {
                let keynote_slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get slides: {}", e)))?;
                Ok(keynote_slides.into_iter().map(Slide::Keynote).collect())
            },
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
            #[cfg(feature = "ole")]
            PresentationImpl::Ppt(_) => Ok(None),
            #[cfg(feature = "ooxml")]
            PresentationImpl::Pptx(pres) => pres.slide_width().map_err(Error::from),
            #[cfg(feature = "iwa")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
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
            #[cfg(feature = "ole")]
            PresentationImpl::Ppt(_) => Ok(None),
            #[cfg(feature = "ooxml")]
            PresentationImpl::Pptx(pres) => pres.slide_height().map_err(Error::from),
            #[cfg(feature = "iwa")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
        }
    }

    /// Extract presentation metadata.
    ///
    /// Returns document properties like title, author, creation date, etc.
    /// The availability of metadata depends on the file format and whether
    /// the properties were set when the presentation was created.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::Presentation;
    ///
    /// let pres = Presentation::open("presentation.pptx")?;
    /// if let Some(metadata) = pres.metadata()? {
    ///     if let Some(title) = metadata.title {
    ///         println!("Title: {}", title);
    ///     }
    ///     if let Some(author) = metadata.author {
    ///         println!("Author: {}", author);
    ///     }
    /// }
    /// # Ok::<(), litchi::common::Error>(())
    /// ```
    pub fn metadata(&self) -> Result<Option<crate::common::Metadata>> {
        // Return cached metadata that was extracted during presentation creation
        Ok(self.cached_metadata.clone())
    }
}
