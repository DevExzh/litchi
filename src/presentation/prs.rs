//! PowerPoint presentation implementation.

use super::Slide;
use super::types::{PptSlideData, PptxSlideData, PresentationImpl};
use crate::common::{Error, Result};

#[cfg(feature = "ole")]
use crate::ole;

#[cfg(feature = "ooxml")]
use crate::ooxml;

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
        // Read file into memory and use smart detection for single-pass parsing
        // This is faster than the old approach of detecting first then parsing again
        let bytes = std::fs::read(path.as_ref())?;
        Self::from_bytes(bytes)
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
    /// - **Single-pass parsing**: Format detection reuses the parsed structure (40-60% faster)
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        // Use smart detection to parse only once
        use crate::common::detection::{DetectedFormat, detect_format_smart};

        let detected = detect_format_smart(bytes).ok_or(Error::NotOfficeFile)?;

        match detected {
            #[cfg(feature = "ole")]
            DetectedFormat::Ppt(ole_file) => {
                // OLE file already parsed - reuse it!
                let mut package =
                    Box::new(ole::ppt::Package::from_ole_file(ole_file).map_err(Error::from)?);

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
            #[cfg(feature = "ooxml")]
            DetectedFormat::Pptx(opc_package) => {
                // OPC package already parsed - reuse it!
                let package = Box::new(
                    ooxml::pptx::Package::from_opc_package(opc_package).map_err(Error::from)?,
                );

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
            #[cfg(feature = "iwa")]
            DetectedFormat::Keynote(zip_archive) => {
                let doc = crate::iwa::keynote::KeynoteDocument::from_zip_archive(zip_archive)
                    .map_err(|e| {
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
            #[cfg(feature = "odf")]
            DetectedFormat::Odp(zip_archive) => {
                let doc = crate::odf::Presentation::from_zip_archive(zip_archive).map_err(|e| {
                    Error::ParseError(format!(
                        "Failed to parse ODP presentation from bytes: {}",
                        e
                    ))
                })?;

                Ok(Self {
                    inner: PresentationImpl::Odp(doc),
                    cached_metadata: Some(crate::common::Metadata::default()),
                    #[cfg(feature = "ooxml")]
                    _pptx_package: None,
                })
            },
            // Handle mismatched formats
            #[allow(unreachable_patterns)]
            _ => Err(Error::InvalidFormat(
                "Detected format is not a presentation format or feature not enabled".to_string(),
            )),
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
            #[cfg(feature = "odf")]
            PresentationImpl::Odp(doc) => {
                let mut text = String::new();
                let slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get ODP slides: {}", e)))?;
                for slide in slides {
                    if let Ok(slide_text) = slide.text() {
                        if !text.is_empty() {
                            text.push_str("\n\n");
                        }
                        text.push_str(slide_text.as_ref());
                    }
                }
                Ok(text)
            },
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
            #[cfg(feature = "odf")]
            PresentationImpl::Odp(doc) => doc
                .slide_count()
                .map_err(|e| Error::ParseError(format!("Failed to get ODP slide count: {}", e))),
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
            #[cfg(feature = "odf")]
            PresentationImpl::Odp(doc) => {
                let odp_slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get ODP slides: {}", e)))?;
                Ok(odp_slides.into_iter().map(Slide::Odp).collect())
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
            #[cfg(feature = "odf")]
            PresentationImpl::Odp(_) => Ok(None), // ODP doesn't expose slide dimensions in unified API yet
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
            #[cfg(feature = "odf")]
            PresentationImpl::Odp(_) => Ok(None), // ODP doesn't expose slide dimensions in unified API yet
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
