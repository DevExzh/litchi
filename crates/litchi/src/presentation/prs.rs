//! PowerPoint presentation implementation.

use super::Slide;
use super::types::PresentationImpl;
use litchi_core::{Error, Result};

use crate::detection_smart::DetectedFormat;

#[cfg(feature = "ppt")]
use crate::ppt;
#[cfg(feature = "ppt")]
use litchi_ole_common::property_set::PropertySetReader;

use std::path::Path;
#[cfg(all(feature = "ppt", any(unix, windows)))]
use std::sync::Arc;

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
    /// Cached metadata extracted during presentation creation.
    ///
    /// Metadata is extracted once during `open()` or `from_bytes()` and cached here
    /// for efficient access. This avoids needing mutable access during `metadata()` calls.
    pub(super) cached_metadata: Option<litchi_core::Metadata>,
}

#[cfg(all(test, feature = "odp"))]
mod flat_odp_tests {
    use super::Presentation;
    use crate::detection_smart::{DetectedFormat, detect_format_smart};
    use litchi_core::detection::FileFormat;

    const FLAT_ODP: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.3"
  office:mimetype="application/vnd.oasis.opendocument.presentation">
  <office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame presentation:class="title"><draw:text-box><text:p>Hello flat slides</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body>
</office:document>"#;

    #[test]
    fn flat_odp_detection_and_facade_reading() {
        match detect_format_smart(FLAT_ODP.to_vec()).expect("flat ODP should be detected") {
            DetectedFormat::FlatOdf(FileFormat::Odp, retained) => assert_eq!(retained, FLAT_ODP),
            _ => panic!("flat ODP was not detected as flat OpenDocument presentation"),
        }

        assert!(matches!(
            Presentation::from_bytes(FLAT_ODP.to_vec()),
            Err(litchi_core::Error::Unsupported(_))
        ));
    }
}

#[cfg(all(test, feature = "pptx", feature = "odp"))]
mod ooxml_odf_polyglot_tests {
    use super::Presentation;
    use std::io::{Cursor, Write};

    fn dual_marker_pptx() -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        // Keep the valid ODF local mimetype first, while the remaining
        // entries form a valid minimal OPC/PPTX package.
        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/presentation.xml", options).unwrap();
        writer
            .write_all(
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/slides/slide1.xml", options).unwrap();
        writer
            .write_all(
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#,
            )
            .unwrap();
        writer.finish().unwrap();
        output.into_inner()
    }

    #[cfg(not(feature = "docx"))]
    fn dual_marker_docx() -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        writer.start_file("mimetype", options).unwrap();
        writer
            .write_all(b"application/vnd.oasis.opendocument.presentation")
            .unwrap();
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("word/document.xml", options).unwrap();
        writer
            .write_all(
                br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>"#,
            )
            .unwrap();
        writer.finish().unwrap();
        output.into_inner()
    }

    #[test]
    fn ooxml_first_precedence_survives_an_odf_local_mimetype_marker() {
        let bytes = dual_marker_pptx();
        assert!(matches!(
            crate::detection_smart::detect_format_smart(bytes.clone()),
            Some(crate::detection_smart::DetectedFormat::Pptx(_))
        ));
        let presentation = Presentation::from_bytes(bytes)
            .expect("OOXML-first precedence should select the valid PPTX owner");
        assert_eq!(presentation.slide_count().unwrap(), 1);
    }

    #[cfg(not(feature = "docx"))]
    #[test]
    fn disabled_docx_owner_keeps_smart_precedence() {
        let bytes = dual_marker_docx();
        assert!(crate::detection_smart::detect_format_smart(bytes.clone()).is_none());
        assert!(crate::detection_smart::detected::detect_prepared_odp(bytes.clone()).is_err());

        let error = Presentation::from_bytes(bytes)
            .err()
            .expect("disabled DOCX owner must not fall through to ODP");
        assert_eq!(error.to_string(), "Not a valid Office file");
    }

    #[test]
    fn invalid_odf_body_still_falls_back_to_typed_odp_validation() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_PRESENTATION)
            .unwrap();
        writer
            .add_file("content.xml", b"<not-an-odp-document/>")
            .unwrap();
        let bytes = writer.finish_to_bytes().unwrap();

        let error = Presentation::from_bytes(bytes)
            .err()
            .expect("invalid ODP body must not be accepted as PPTX");
        assert!(error.to_string().contains("ODP"));
    }

    #[test]
    fn valid_odp_wins_after_the_ooxml_probe_fails() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer
            .set_mimetype(litchi_odf_common::constants::ODF_PRESENTATION)
            .unwrap();
        writer
            .add_file(
                "content.xml",
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
            )
            .unwrap();

        assert!(Presentation::from_bytes(writer.finish_to_bytes().unwrap()).is_ok());
    }
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
        #[cfg(all(feature = "ppt", any(unix, windows), not(feature = "pptx")))]
        if let Some(presentation) = Self::open_native_ppt_path(path.as_ref())? {
            return Ok(presentation);
        }

        #[cfg(feature = "pptx")]
        {
            Self::open_with_limits(path, crate::pptx::ReadLimits::default())
        }

        #[cfg(not(feature = "pptx"))]
        {
            let bytes = std::fs::read(path.as_ref())?;
            Self::from_bytes(bytes)
        }
    }

    /// Open a presentation with an explicit PPTX/OPC resource policy.
    ///
    /// The policy applies to OOXML-suffixed paths and ZIP-magic candidates.
    /// Legacy PowerPoint, Keynote, and OpenDocument inputs continue through
    /// their native readers.
    #[cfg(feature = "pptx")]
    pub fn open_with_limits<P: AsRef<Path>>(
        path: P,
        limits: crate::pptx::ReadLimits,
    ) -> Result<Self> {
        if let Some(detected) =
            crate::detection_smart::detected::detect_ooxml_path_with_limits(path.as_ref(), limits)
                .map_err(crate::map_ooxml_error)?
        {
            return Self::from_detected(detected);
        }

        #[cfg(all(feature = "ppt", any(unix, windows)))]
        if let Some(presentation) = Self::open_native_ppt_path(path.as_ref())? {
            return Ok(presentation);
        }

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
    /// - OLE2 and OOXML detection return parsed owners that their loaders reuse
    /// - Other detection results retain the moved buffer for loaders that may parse it afterward
    /// - Ideal for network data, streams, or in-memory content
    /// - No temporary files created
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        #[cfg(feature = "odp")]
        let bytes = match crate::detection_smart::detected::detect_prepared_odp(bytes) {
            Ok(prepared) => {
                let doc = litchi_odp::Presentation::from_prepared_package(prepared)?;
                return Ok(Self {
                    inner: PresentationImpl::Odp(doc),
                    cached_metadata: Some(litchi_core::Metadata::default()),
                });
            },
            Err(bytes) => bytes,
        };

        #[cfg(feature = "pptx")]
        {
            Self::from_bytes_with_limits(bytes, crate::pptx::ReadLimits::default())
        }

        #[cfg(not(feature = "pptx"))]
        {
            let detected =
                crate::detection_smart::detect_format_smart(bytes).ok_or(Error::NotOfficeFile)?;
            Self::from_detected(detected)
        }
    }

    #[cfg(all(feature = "ppt", any(unix, windows)))]
    fn open_native_ppt_path(path: &Path) -> Result<Option<Self>> {
        let source = Arc::new(litchi_core::FileSource::open(path)?);
        let shared = match litchi_cfb::SharedOleFile::open(source) {
            Ok(shared) => shared,
            // Classification failures precede PPT ownership. Preserve the
            // established byte/detection fallback for malformed or non-CFB
            // inputs instead of leaking a PPT-specific error.
            Err(_error) => return Ok(None),
        };

        #[cfg(feature = "doc")]
        if shared.exists(&["WordDocument"]) {
            // Match the existing smart detector's DOC-before-PPT precedence
            // for valid OLE polyglots. The byte fallback below then returns
            // the established non-presentation result for the DOC owner.
            return Ok(None);
        }

        let Some(package) =
            ppt::SourceBackedPackage::from_shared_if_powerpoint(shared).map_err(Error::from)?
        else {
            return Ok(None);
        };

        let cached_metadata = package
            .metadata()
            .ok()
            .map(litchi_core::Metadata::from)
            .filter(|metadata| metadata.has_data());
        let pres = package.presentation().map_err(Error::from)?;
        Ok(Some(Self {
            inner: PresentationImpl::Ppt(pres),
            cached_metadata,
        }))
    }

    /// Create a presentation from bytes with an explicit PPTX/OPC resource
    /// policy. The policy is consulted only while probing an OOXML ZIP candidate.
    #[cfg(feature = "pptx")]
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: crate::pptx::ReadLimits) -> Result<Self> {
        let detected = crate::detection_smart::detect_format_smart_with_limits(bytes, limits)
            .ok_or(Error::NotOfficeFile)?;
        Self::from_detected(detected)
    }

    fn from_detected(detected: crate::detection_smart::DetectedFormat) -> Result<Self> {
        match detected {
            #[cfg(feature = "ppt")]
            DetectedFormat::Ppt(ole_file) => {
                // OLE file already parsed - reuse it!
                let mut package =
                    Box::new(ppt::Package::from_ole_file(ole_file).map_err(Error::from)?);

                // Extract metadata from OLE property streams
                let cached_metadata = package
                    .ole_file()
                    .get_metadata()
                    .ok()
                    .map(|ole_metadata| {
                        let metadata: litchi_core::Metadata = ole_metadata.into();
                        metadata
                    })
                    .filter(|metadata| metadata.has_data());

                let pres = package.presentation().map_err(Error::from)?;

                Ok(Self {
                    inner: PresentationImpl::Ppt(pres),
                    cached_metadata,
                })
            },
            #[cfg(feature = "pptx")]
            DetectedFormat::Pptx(opc_package) => {
                // OPC package already parsed - reuse it!
                let cached_metadata = crate::ooxml_common::properties::read(&opc_package)
                    .map_err(crate::map_ooxml_error)?
                    .map(litchi_core::Metadata::from)
                    .filter(|metadata| metadata.has_data());
                let package = Box::new(
                    crate::pptx::Package::from_opc_package(opc_package)
                        .map_err(crate::map_ooxml_error)?,
                );

                Ok(Self {
                    inner: PresentationImpl::Pptx(package),
                    cached_metadata,
                })
            },
            #[cfg(feature = "keynote")]
            DetectedFormat::Keynote(data) => {
                let doc = litchi_keynote::Package::from_bytes(&data).map_err(|e| {
                    Error::ParseError(format!("Failed to open Keynote from bytes: {}", e))
                })?;

                // Extract Keynote metadata from bundle properties
                let cached_metadata = doc
                    .metadata()
                    .ok()
                    .flatten()
                    .filter(|metadata| metadata.has_data());

                Ok(Self {
                    inner: PresentationImpl::Keynote(doc),
                    cached_metadata,
                })
            },
            #[cfg(feature = "odp")]
            DetectedFormat::FlatOdf(format, data) => {
                let _ = data;
                Err(Error::Unsupported(format!(
                    "flat OpenDocument {:?} is detected but the dedicated family facade exposes packaged parsing only",
                    format
                )))
            },
            #[cfg(feature = "odp")]
            DetectedFormat::Odp(data) => {
                let doc = litchi_odp::Presentation::from_bytes(data).map_err(|e| {
                    Error::ParseError(format!(
                        "Failed to parse ODP presentation from bytes: {}",
                        e
                    ))
                })?;

                Ok(Self {
                    inner: PresentationImpl::Odp(doc),
                    cached_metadata: Some(litchi_core::Metadata::default()),
                })
            },
            // Handle mismatched formats
            #[allow(
                unreachable_patterns,
                reason = "match arms are feature-gated; the fallback is unreachable when every format feature is enabled"
            )]
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
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(pres) => pres.text().map_err(Error::from),
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => {
                // PPTX presentations need to extract text from all slides
                let pres = package.presentation().map_err(crate::map_ooxml_error)?;
                let slides = pres.slides().map_err(crate::map_ooxml_error)?;
                let mut texts = Vec::new();
                for slide in slides {
                    let text = slide.text().map_err(crate::map_ooxml_error)?;
                    if !text.is_empty() {
                        texts.push(text);
                    }
                }
                Ok(texts.join("\n\n"))
            },
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(doc) => doc.text().map_err(|e| {
                Error::ParseError(format!("Failed to extract text from Keynote: {}", e))
            }),
            #[cfg(feature = "odp")]
            PresentationImpl::Odp(doc) => doc
                .text()
                .map_err(|e| Error::ParseError(format!("Failed to extract ODP text: {}", e))),
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
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(pres) => Ok(pres.slide_count()),
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => package
                .presentation()
                .map_err(crate::map_ooxml_error)?
                .slide_count()
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(doc) => {
                let slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get slides: {}", e)))?;
                Ok(slides.len())
            },
            #[cfg(feature = "odp")]
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
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(pres) => {
                use super::types::LegacySlideData;
                // Extract slide data to avoid lifetime issues
                let ppt_slides = pres.slides().map_err(Error::from)?;
                ppt_slides
                    .iter()
                    .map(|s| {
                        let text = s.text().map_err(Error::from)?.to_string();
                        let slide_number = s.slide_number();
                        let shape_count = s.shape_count().unwrap_or(0);
                        Ok(Slide::Ppt(LegacySlideData {
                            text,
                            slide_number,
                            shape_count,
                        }))
                    })
                    .collect()
            },
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => {
                use super::types::SlideData;
                let pres = package.presentation().map_err(crate::map_ooxml_error)?;
                let slides = pres.slides().map_err(crate::map_ooxml_error)?;
                // Extract slide data immediately to avoid lifetime issues
                slides
                    .iter()
                    .map(|s| {
                        let text = s.text().map_err(crate::map_ooxml_error)?;
                        let name = Some(s.name().map_err(crate::map_ooxml_error)?);
                        Ok(Slide::Pptx(SlideData { text, name }))
                    })
                    .collect()
            },
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(doc) => {
                let keynote_slides = doc
                    .slides()
                    .map_err(|e| Error::ParseError(format!("Failed to get slides: {}", e)))?;
                Ok(keynote_slides
                    .iter()
                    .enumerate()
                    .map(|(index, slide)| {
                        let name = slide.name().map(str::to_owned);
                        let title = slide.title().map(str::to_owned);
                        let text = slide.plain_text();
                        Slide::Keynote {
                            number: index + 1,
                            name,
                            title,
                            text,
                        }
                    })
                    .collect())
            },
            #[cfg(feature = "odp")]
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
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(_) => Ok(None),
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => package
                .presentation()
                .map_err(crate::map_ooxml_error)?
                .slide_size()
                .map(|(width, _)| Some(width))
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
            #[cfg(feature = "odp")]
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
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(_) => Ok(None),
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => package
                .presentation()
                .map_err(crate::map_ooxml_error)?
                .slide_size()
                .map(|(_, height)| Some(height))
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
            #[cfg(feature = "odp")]
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
    pub fn metadata(&self) -> Result<Option<litchi_core::Metadata>> {
        // Return cached metadata that was extracted during presentation creation
        Ok(self.cached_metadata.clone())
    }

    /// Fast text extraction for markdown conversion (internal use).
    ///
    /// This method is optimized for PPT files by skipping shape parsing.
    /// For other formats, it falls back to the standard text extraction.
    ///
    /// # Performance
    ///
    /// - PPT: Uses fast path that bypasses shape parsing (3-10x faster)
    /// - PPTX/Keynote/ODP: Falls back to standard extraction
    ///
    /// # Returns
    ///
    /// Vector of (slide_number, text) tuples for each slide
    #[doc(hidden)]
    pub fn extract_text_for_markdown(&self) -> Result<Vec<(usize, String)>> {
        // Fast PPT path when only `ppt` is enabled. In this configuration
        // PresentationImpl can only be Ppt, so we destructure directly and
        // return early.
        #[cfg(all(
            feature = "ppt",
            not(any(feature = "pptx", feature = "keynote", feature = "odp"))
        ))]
        {
            let PresentationImpl::Ppt(pres) = &self.inner;
            pres.extract_text_fast().map_err(Error::from)
        }

        // When multiple presentation formats are compiled in, prefer the fast
        // PPT extractor but keep the generic slide-based fallback for other
        // formats.
        #[cfg(not(all(
            feature = "ppt",
            not(any(feature = "pptx", feature = "keynote", feature = "odp"))
        )))]
        {
            #[cfg(feature = "ppt")]
            if let PresentationImpl::Ppt(pres) = &self.inner {
                return pres.extract_text_fast().map_err(Error::from);
            }

            // For other formats, extract from slides (slower but works)
            let slides = self.slides()?;
            slides
                .iter()
                .enumerate()
                .map(|(idx, slide)| slide.text().map(|text| (idx + 1, text.to_string())))
                .collect()
        }
    }
}

#[cfg(all(test, feature = "keynote"))]
mod keynote_facade_tests {
    use super::Presentation;
    use crate::presentation::Slide;
    use std::path::PathBuf;

    fn fixture_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key")
    }

    #[test]
    fn slides_preserve_keynote_identity_and_plain_text_semantics() {
        let path = fixture_path();
        let native = litchi_keynote::Package::open(&path).expect("open native Keynote fixture");
        let native_slides = native.slides().expect("read native Keynote slides");
        let facade = Presentation::open(&path).expect("open Keynote through presentation facade");
        let facade_slides = facade.slides().expect("read facade slides");

        assert_eq!(facade_slides.len(), native_slides.len());
        for (index, (facade_slide, native_slide)) in
            facade_slides.iter().zip(native_slides).enumerate()
        {
            let Slide::Keynote {
                number,
                name,
                title,
                text,
            } = facade_slide
            else {
                panic!("Keynote presentation yielded a non-Keynote facade slide")
            };

            assert_eq!(*number, index + 1);
            assert_eq!(name.as_deref(), native_slide.name());
            assert_eq!(title.as_deref(), native_slide.title());
            assert_eq!(text, &native_slide.plain_text());
        }
    }
}

#[cfg(all(test, feature = "pptx", feature = "ppt"))]
mod tests {
    use super::*;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_open_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path);
        assert!(pres.is_ok(), "Failed to open PPTX file: {:?}", pres.err());
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_open_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path);
        assert!(pres.is_ok(), "Failed to open PPT file: {:?}", pres.err());
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_from_bytes_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let bytes = std::fs::read(&path).expect("Failed to read file");
        let pres = Presentation::from_bytes(bytes);
        assert!(
            pres.is_ok(),
            "Failed to load PPTX from bytes: {:?}",
            pres.err()
        );
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_from_bytes_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let bytes = std::fs::read(&path).expect("Failed to read file");
        let pres = Presentation::from_bytes(bytes);
        assert!(
            pres.is_ok(),
            "Failed to load PPT from bytes: {:?}",
            pres.err()
        );
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_text_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let _text = pres.text().expect("Failed to extract text");
        // Text may be empty for some presentations
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_text_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let _text = pres.text().expect("Failed to extract text");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_slide_count_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let count = pres.slide_count().expect("Failed to get slide count");
        assert!(count > 0, "Expected at least one slide");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_slide_count_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let count = pres.slide_count().expect("Failed to get slide count");
        assert!(count > 0, "Expected at least one slide");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_slides_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let slides = pres.slides().expect("Failed to get slides");
        assert!(!slides.is_empty(), "Expected at least one slide");

        // Test that we can access text from slides
        for slide in slides {
            let _text = slide.text().expect("Failed to get slide text");
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_slides_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides = pres.slides().expect("Failed to get slides");
        assert!(!slides.is_empty(), "Expected at least one slide");

        for slide in slides {
            let _text = slide.text().expect("Failed to get slide text");
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_slide_dimensions_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let _width = pres.slide_width().expect("Failed to get slide width");
        let _height = pres.slide_height().expect("Failed to get slide height");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_metadata_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let metadata = pres.metadata().expect("Failed to get metadata");
        // Metadata may or may not be present
        let _ = metadata.map(|m| m.title);
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_metadata_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let metadata = pres.metadata().expect("Failed to get metadata");
        let _ = metadata.map(|m| m.title);
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_open_nonexistent_file() {
        let path = test_data_path().join("nonexistent_file.pptx");
        let result = Presentation::open(&path);
        assert!(result.is_err(), "Expected error for nonexistent file");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_from_bytes_invalid_data() {
        let bytes = b"This is not a valid presentation file".to_vec();
        let result = Presentation::from_bytes(bytes);
        assert!(result.is_err(), "Expected error for invalid data");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_charts_pptx() {
        // Test presentations with various chart types
        let chart_files = [
            "ooxml/pptx/line-chart.pptx",
            "ooxml/pptx/pie-chart.pptx",
            "ooxml/pptx/radar-chart.pptx",
            "ooxml/pptx/scatter-chart.pptx",
        ];

        for file in &chart_files {
            let path = test_data_path().join(file);
            if path.exists() {
                let pres = Presentation::open(&path);
                assert!(pres.is_ok(), "Failed to open {}", file);

                if let Ok(p) = pres {
                    let count = p.slide_count().expect("Failed to get slide count");
                    assert!(count > 0, "Expected at least one slide in {}", file);
                }
            }
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_connectors_pptx() {
        let test_files = [
            "ooxml/pptx/connectorConnection.pptx",
            "ooxml/pptx/curvedConnectors.pptx",
            "ooxml/pptx/elbowConnectors.pptx",
        ];

        for file in &test_files {
            let path = test_data_path().join(file);
            if path.exists() {
                let pres = Presentation::open(&path);
                assert!(pres.is_ok(), "Failed to open {}", file);
            }
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_text_shapes_ppt() {
        // Use SampleShow.ppt to avoid metadata overflow issues in some test files
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides = pres.slides().expect("Failed to get slides");
        assert!(!slides.is_empty(), "Expected at least one slide");

        for slide in slides {
            // shape_count is only available for PPT format
            let _ = slide.shape_count();
        }
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_extract_text_for_markdown_ppt() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let pres = Presentation::open(&path).expect("Failed to open PPT");
        let slides_text = pres
            .extract_text_for_markdown()
            .expect("Failed to extract text");
        // Verify we got text for each slide
        assert!(!slides_text.is_empty(), "Expected text extraction results");
    }

    #[test]
    #[cfg(all(feature = "pptx", feature = "ppt"))]
    fn test_presentation_extract_text_for_markdown_pptx() {
        let path = test_data_path().join("ooxml/pptx/sample.pptx");
        let pres = Presentation::open(&path).expect("Failed to open PPTX");
        let slides_text = pres
            .extract_text_for_markdown()
            .expect("Failed to extract text");
        assert!(!slides_text.is_empty(), "Expected text extraction results");
    }
}

#[cfg(all(test, feature = "ppt", any(unix, windows)))]
mod native_ppt_path_tests {
    use super::Presentation;
    use litchi_cfb::{OleFile, OleWriter};
    use std::io::Cursor;
    use std::path::PathBuf;

    fn test_data_path() -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data")
    }

    #[test]
    fn native_path_matches_byte_facade_for_core_queries_and_metadata() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let from_path = Presentation::open(&path).expect("open native PPT path");
        let from_bytes = Presentation::from_bytes(std::fs::read(&path).expect("read PPT"))
            .expect("open PPT bytes");

        assert_eq!(
            from_path.slide_count().unwrap(),
            from_bytes.slide_count().unwrap()
        );
        assert_eq!(from_path.text().unwrap(), from_bytes.text().unwrap());
        assert_eq!(
            from_path
                .metadata()
                .unwrap()
                .map(|metadata| format!("{metadata:?}")),
            from_bytes
                .metadata()
                .unwrap()
                .map(|metadata| format!("{metadata:?}"))
        );
    }

    #[test]
    fn non_powerpoint_ole_keeps_the_existing_facade_fallback() {
        let path = test_data_path().join("ole/doc/documentProperties.doc");
        let path_error = match Presentation::open(&path) {
            Ok(_) => panic!("a Word OLE package is not a presentation"),
            Err(error) => error.to_string(),
        };
        let bytes_error = match Presentation::from_bytes(std::fs::read(&path).expect("read DOC")) {
            Ok(_) => panic!("a Word OLE package is not a presentation"),
            Err(error) => error.to_string(),
        };

        assert_eq!(path_error, bytes_error);
    }

    #[test]
    fn malformed_ole_keeps_the_existing_facade_fallback() {
        let mut bytes = litchi_core::detection::utils::OLE2_SIGNATURE.to_vec();
        bytes.resize(512, 0);
        let temporary = tempfile::NamedTempFile::new().expect("temporary malformed OLE path");
        std::fs::write(temporary.path(), &bytes).expect("write malformed OLE");

        let path_error = match Presentation::open(temporary.path()) {
            Ok(_) => panic!("malformed OLE path must not become a presentation"),
            Err(error) => error.to_string(),
        };
        let bytes_error = match Presentation::from_bytes(bytes) {
            Ok(_) => panic!("malformed OLE bytes must not become a presentation"),
            Err(error) => error.to_string(),
        };
        assert_eq!(path_error, bytes_error);
    }

    #[cfg(feature = "pptx")]
    #[test]
    fn ooxml_suffix_preflight_precedes_native_ole_routing() {
        let source = test_data_path().join("ole/ppt/SampleShow.ppt");
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary OOXML-suffixed path");
        std::fs::write(
            temporary.path(),
            std::fs::read(source).expect("read native PPT"),
        )
        .expect("write OOXML-suffixed PPT");

        let error = match Presentation::open(temporary.path()) {
            Ok(_) => panic!("OOXML suffix must reject native OLE bytes"),
            Err(error) => error.to_string(),
        };
        assert!(
            error.contains("OOXML-suffixed input does not have ZIP magic"),
            "unexpected extension error: {error}"
        );
    }

    #[cfg(feature = "doc")]
    #[test]
    fn doc_host_stream_keeps_doc_before_ppt_precedence() {
        let doc_path = test_data_path().join("ole/doc/documentProperties.doc");
        let ppt_path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let mut doc = OleFile::open(Cursor::new(std::fs::read(doc_path).expect("read DOC")))
            .expect("open DOC OLE");
        let mut ppt = OleFile::open(Cursor::new(std::fs::read(ppt_path).expect("read PPT")))
            .expect("open PPT OLE");
        let mut writer = OleWriter::new();
        writer
            .create_stream(
                &["WordDocument"],
                &doc.open_stream(&["WordDocument"]).expect("DOC stream"),
            )
            .expect("write DOC host stream");
        writer
            .create_stream(
                &["PowerPoint Document"],
                &ppt.open_stream(&["PowerPoint Document"])
                    .expect("PPT stream"),
            )
            .expect("write PPT document stream");
        writer
            .create_stream(
                &["Current User"],
                &ppt.open_stream(&["Current User"])
                    .expect("PPT current-user stream"),
            )
            .expect("write PPT current-user stream");
        let mut output = Cursor::new(Vec::new());
        writer
            .write_to(&mut output)
            .expect("serialize OLE polyglot");
        let bytes = output.into_inner();

        let bytes_error = match Presentation::from_bytes(bytes.clone()) {
            Ok(_) => panic!("DOC precedence should reject the polyglot as a presentation"),
            Err(error) => error.to_string(),
        };
        let temporary = tempfile::NamedTempFile::new().expect("temporary OLE path");
        std::fs::write(temporary.path(), bytes).expect("write OLE polyglot");
        let path_error = match Presentation::open(temporary.path()) {
            Ok(_) => panic!("DOC precedence should reject the polyglot as a presentation"),
            Err(error) => error.to_string(),
        };

        assert_eq!(path_error, bytes_error);
    }
}
