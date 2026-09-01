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
use std::sync::OnceLock;

/// Detached point-in-time snapshot of one PPTX slide's catalog metadata.
///
/// The position is zero-based and the slide ID is the producer-visible
/// `p:sldId@id` value retained by the source catalog. Constructing a
/// descriptor never reads slide XML or related media payloads.
#[cfg(feature = "pptx")]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideDescriptor {
    position: usize,
    slide_id: u32,
}

#[cfg(feature = "pptx")]
impl SlideDescriptor {
    fn new(position: usize, slide_id: u32) -> Self {
        Self { position, slide_id }
    }

    /// Zero-based position in the presentation's ordered slide list.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Stable producer-visible `p:sldId@id` value.
    #[must_use]
    pub const fn slide_id(&self) -> u32 {
        self.slide_id
    }
}

#[cfg(feature = "pptx")]
fn map_pptx_catalog_error(error: crate::pptx::Error) -> Error {
    match error {
        crate::pptx::Error::Opc(crate::opc::OpcError::SourceChanged { expected, actual }) => {
            Error::SourceChanged {
                expected,
                observed: actual,
            }
        },
        crate::pptx::Error::Opc(crate::opc::OpcError::Cancelled) => {
            Error::Other("PPTX source operation cancelled".to_owned())
        },
        other => crate::map_ooxml_error(other),
    }
}

#[cfg(feature = "pptx")]
fn reserve_slide_catalog(capacity: usize) -> Result<Vec<SlideDescriptor>> {
    let mut descriptors = Vec::new();
    descriptors
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "PPTX slide descriptors",
            source,
        })?;
    Ok(descriptors)
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
    pub(super) inner: PresentationImpl,
    /// Metadata cache initialized eagerly for materialized owners and lazily
    /// for the source-backed PPTX owner. A source read failure leaves the
    /// cell unset so a later call can retry after the source is restored.
    pub(super) cached_metadata: OnceLock<Option<litchi_core::Metadata>>,
}

#[cfg(all(feature = "odp", any(unix, windows)))]
fn map_odp_error(error: Error, _operation: &str) -> Error {
    match error {
        error @ (Error::Io(_)
        | Error::Allocation { .. }
        | Error::ResourceLimit(_)
        | Error::SourceChanged { .. }
        | Error::InvalidFormat(_)
        | Error::ParseError(_)) => error,
        other => other,
    }
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

#[cfg(all(
    test,
    feature = "odp",
    any(unix, windows),
    any(
        feature = "pptx",
        not(any(feature = "docx", feature = "xlsx", feature = "xlsb"))
    )
))]
mod source_odp_path_tests {
    use super::{Presentation, PresentationImpl};
    use litchi_core::Error;

    const MIME: &str = "application/vnd.oasis.opendocument.presentation";
    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn package() -> Vec<u8> {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:text="{TEXT}"><office:body><office:presentation><draw:page draw:name="one"><draw:frame><draw:text-box><text:p>root source</text:p></draw:text-box></draw:frame></draw:page><draw:page draw:name="two"><draw:frame><draw:text-box><text:p>second slide</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
        );
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).expect("ODP MIME");
        writer
            .add_file("content.xml", content.as_bytes())
            .expect("ODP content");
        writer.finish_to_bytes().expect("ODP package")
    }

    #[test]
    fn path_uses_odp_source_owner_and_matches_eager_projection() {
        let bytes = package();
        let temporary = tempfile::NamedTempFile::new().expect("temporary ODP path");
        std::fs::write(temporary.path(), &bytes).expect("write ODP fixture");

        let source = Presentation::open(temporary.path()).expect("source-backed ODP");
        assert!(matches!(&source.inner, PresentationImpl::OdpSource(_)));
        let eager = Presentation::from_bytes(bytes).expect("eager ODP control");

        assert_eq!(source.slide_count().unwrap(), eager.slide_count().unwrap());
        assert_eq!(source.text().unwrap(), eager.text().unwrap());
        let source_slide_text = source
            .slides()
            .unwrap()
            .into_iter()
            .map(|slide| slide.text().unwrap())
            .collect::<Vec<_>>();
        let eager_slide_text = eager
            .slides()
            .unwrap()
            .into_iter()
            .map(|slide| slide.text().unwrap())
            .collect::<Vec<_>>();
        assert_eq!(source_slide_text, eager_slide_text);
        assert!(source.metadata().unwrap().is_some());
        assert_eq!(
            source.slide(0).unwrap().unwrap().text().unwrap(),
            eager.slide(0).unwrap().unwrap().text().unwrap()
        );
        assert_eq!(source.slide_width().unwrap(), eager.slide_width().unwrap());
        assert_eq!(
            source.slide_height().unwrap(),
            eager.slide_height().unwrap()
        );
    }

    #[cfg(feature = "pptx")]
    #[test]
    fn canonical_and_renamed_odp_paths_use_source_owner_under_tight_pptx_limit() {
        let bytes = package();
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("consistent input limit");

        for suffix in [".odp", ".pptx"] {
            let temporary = tempfile::Builder::new()
                .suffix(suffix)
                .tempfile()
                .expect("temporary ODP path");
            std::fs::write(temporary.path(), &bytes).expect("write ODP fixture");
            let presentation = Presentation::open_with_limits(temporary.path(), limits)
                .expect("ordinary ODP must use its source owner before PPTX limits");
            assert!(matches!(
                &presentation.inner,
                PresentationImpl::OdpSource(_)
            ));
        }
    }

    #[test]
    fn path_source_metadata_and_dimensions_report_stale_source() {
        let temporary = tempfile::NamedTempFile::new().expect("temporary ODP path");
        std::fs::write(temporary.path(), package()).expect("write ODP fixture");
        let presentation = Presentation::open(temporary.path()).expect("source-backed ODP");
        assert!(presentation.metadata().unwrap().is_some());

        let mut file = std::fs::OpenOptions::new()
            .append(true)
            .open(temporary.path())
            .expect("reopen ODP for mutation");
        use std::io::Write;
        file.write_all(b"source mutation")
            .expect("mutate ODP source");

        assert!(matches!(
            presentation.metadata(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.slide_width(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.slide_height(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.slide_count(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.slides(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.slide(0),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.text(),
            Err(Error::SourceChanged { .. })
        ));
        assert!(matches!(
            presentation.extract_text_for_markdown(),
            Err(Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn valid_mime_with_malformed_odp_body_stays_with_odp_owner_error() {
        let mut writer = litchi_odf_common::core::PackageWriter::new();
        writer.set_mimetype(MIME).expect("ODP MIME");
        writer
            .add_file("content.xml", b"<not-an-odp-document/>")
            .expect("malformed content");
        let bytes = writer.finish_to_bytes().expect("ODP package");
        let temporary = tempfile::NamedTempFile::new().expect("temporary ODP path");
        std::fs::write(temporary.path(), bytes).expect("write ODP fixture");

        let error = match Presentation::open(temporary.path()) {
            Ok(_) => panic!("malformed ODP must fail at the source owner"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            Error::InvalidFormat(_) | Error::ParseError(_)
        ));
    }
}

#[cfg(all(
    test,
    not(feature = "pptx"),
    any(feature = "ppt", feature = "odp", feature = "keynote"),
    any(unix, windows)
))]
mod no_pptx_path_tests {
    use super::Presentation;

    #[test]
    fn generic_path_is_bounded_by_neutral_ceiling_without_pptx() {
        let temporary = tempfile::NamedTempFile::new().expect("temporary sparse input");
        let neutral =
            crate::detection_smart::detected::UNIFIED_PRESENTATION_FALLBACK_MAX_INPUT_BYTES;
        let neutral_plus_one = neutral + 1;
        temporary
            .as_file()
            .set_len(neutral_plus_one)
            .expect("create sparse input");

        let error = match crate::detection_smart::detected::detect_presentation_source_path(
            temporary.path(),
        ) {
            Ok(_) => panic!("generic input over the neutral ceiling must fail"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            litchi_core::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::InputBytes
                    && limit.observed == neutral_plus_one
                    && limit.limit == neutral
        ));

        let public_error = match Presentation::open(temporary.path()) {
            Ok(_) => panic!("generic input over the neutral ceiling must fail publicly"),
            Err(error) => error,
        };
        assert!(matches!(
            public_error,
            litchi_core::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::InputBytes
                    && limit.observed == neutral_plus_one
                    && limit.limit == neutral
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
        writer.start_file("META-INF/manifest.xml", options).unwrap();
        writer
            .write_all(
                br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.presentation"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#,
            )
            .unwrap();
        writer.start_file("content.xml", options).unwrap();
        writer
            .write_all(
                br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#,
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

    #[test]
    fn path_ooxml_first_precedence_survives_an_odf_local_mimetype_marker() {
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), dual_marker_pptx()).expect("write PPTX polyglot");
        let presentation = Presentation::open(temporary.path())
            .expect("path OOXML-first precedence should select PPTX");
        assert_eq!(presentation.slide_count().unwrap(), 1);
        assert!(presentation.slide(0).unwrap().is_some());
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

    #[cfg(all(not(feature = "docx"), any(unix, windows)))]
    #[test]
    fn path_disabled_docx_owner_cannot_fall_through_to_valid_odp_marker() {
        let temporary = tempfile::Builder::new()
            .suffix(".odp")
            .tempfile()
            .expect("temporary OOXML/ODP polyglot path");
        std::fs::write(temporary.path(), dual_marker_docx()).expect("write OOXML/ODP polyglot");

        let error = Presentation::open(temporary.path())
            .err()
            .expect("disabled DOCX owner must retain OOXML precedence");
        assert!(matches!(error, litchi_core::Error::NotOfficeFile));
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
        #[cfg(feature = "pptx")]
        {
            Self::open_with_limits(path, crate::pptx::ReadLimits::default())
        }

        #[cfg(all(
            not(feature = "pptx"),
            any(feature = "ppt", feature = "odp", feature = "keynote")
        ))]
        {
            match crate::detection_smart::detected::detect_presentation_source_path(path.as_ref())?
            {
                #[cfg(feature = "odp")]
                crate::detection_smart::detected::PresentationSourcePathDetection::Odp(
                    presentation,
                ) => Ok(Self {
                    inner: PresentationImpl::OdpSource(presentation),
                    cached_metadata: OnceLock::from(Some(litchi_core::Metadata::default())),
                }),
                #[cfg(feature = "ppt")]
                crate::detection_smart::detected::PresentationSourcePathDetection::Ppt(package) => {
                    Self::from_native_ppt_package(package)
                },
                crate::detection_smart::detected::PresentationSourcePathDetection::Bytes(bytes) => {
                    Self::from_bytes(bytes)
                },
            }
        }

        #[cfg(all(
            not(feature = "pptx"),
            not(any(feature = "ppt", feature = "odp", feature = "keynote"))
        ))]
        {
            let _ = path;
            Err(Error::NotOfficeFile)
        }
    }

    #[cfg(feature = "pptx")]
    fn map_source_opc_error(error: crate::opc::OpcError) -> Error {
        match error {
            crate::opc::OpcError::SourceChanged { expected, actual } => Error::SourceChanged {
                expected,
                observed: actual,
            },
            crate::opc::OpcError::ReadLimit {
                resource,
                actual,
                maximum,
            } => Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: Self::map_opc_read_resource(resource),
                observed: actual,
                limit: maximum,
                scope: format!("OPC {resource}").into(),
            }),
            other => other.into(),
        }
    }

    #[cfg(feature = "pptx")]
    fn map_opc_read_resource(resource: crate::opc::ReadResource) -> litchi_core::Resource {
        match resource {
            crate::opc::ReadResource::InputBytes
            | crate::opc::ReadResource::ArchiveMemberNameBytes
            | crate::opc::ReadResource::ArchiveMetadataBytes
            | crate::opc::ReadResource::ArchiveCompressedBytes
            | crate::opc::ReadResource::ArchiveEntryBytes
            | crate::opc::ReadResource::ArchiveTotalBytes
            | crate::opc::ReadResource::PartBytes
            | crate::opc::ReadResource::TotalPartBytes
            | crate::opc::ReadResource::ContentTypesBytes
            | crate::opc::ReadResource::RelationshipXmlBytes
            | crate::opc::ReadResource::TotalRelationshipXmlBytes
            | crate::opc::ReadResource::XmlAttributeBytes
            | crate::opc::ReadResource::RelationshipTargetBytes => {
                litchi_core::Resource::InputBytes
            },
            crate::opc::ReadResource::ArchiveMembers
            | crate::opc::ReadResource::Parts
            | crate::opc::ReadResource::ContentTypeMappings
            | crate::opc::ReadResource::RelationshipParts
            | crate::opc::ReadResource::RelationshipsPerPart
            | crate::opc::ReadResource::TotalRelationships
            | crate::opc::ReadResource::RelationshipGraphNodes => litchi_core::Resource::Objects,
            crate::opc::ReadResource::XmlEvents
            | crate::opc::ReadResource::TotalRelationshipXmlEvents => litchi_core::Resource::Work,
            crate::opc::ReadResource::XmlDepth => litchi_core::Resource::Depth,
            _ => litchi_core::Resource::Work,
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
        #[cfg(any(unix, windows))]
        {
            let detected = crate::detection_smart::detected::detect_pptx_source_path_with_limits(
                path.as_ref(),
                limits,
            )
            .map_err(|error| match error {
                crate::detection_smart::detected::PptxSourcePathError::Opc(error) => {
                    Self::map_source_opc_error(error)
                },
                crate::detection_smart::detected::PptxSourcePathError::Pptx(error) => {
                    map_pptx_catalog_error(error)
                },
                #[cfg(any(feature = "odp", feature = "ppt"))]
                crate::detection_smart::detected::PptxSourcePathError::Source(error) => error,
            })?;
            let detected = detected.ok_or(Error::NotOfficeFile)?;
            return match detected {
                crate::detection_smart::detected::PptxSourcePathDetection::Pptx(presentation) => {
                    Self::from_source_backed_pptx(presentation)
                },
                #[cfg(feature = "odp")]
                crate::detection_smart::detected::PptxSourcePathDetection::Odp(presentation) => {
                    Ok(Self {
                        inner: PresentationImpl::OdpSource(presentation),
                        cached_metadata: OnceLock::from(Some(litchi_core::Metadata::default())),
                    })
                },
                #[cfg(feature = "ppt")]
                crate::detection_smart::detected::PptxSourcePathDetection::Ppt(package) => {
                    Self::from_native_ppt_package(package)
                },
                crate::detection_smart::detected::PptxSourcePathDetection::OtherOoxml(format) => {
                    let _ = format;
                    Err(Error::InvalidFormat(
                        "Detected format is not a presentation format or feature not enabled"
                            .to_owned(),
                    ))
                },
                crate::detection_smart::detected::PptxSourcePathDetection::DisabledOtherOoxml(
                    format,
                ) => {
                    let _ = format;
                    Err(Error::NotOfficeFile)
                },
                crate::detection_smart::detected::PptxSourcePathDetection::Bytes(bytes) => {
                    Self::from_bytes_with_limits(bytes, limits)
                },
            };
        }

        #[cfg(not(any(unix, windows)))]
        {
            let bytes = crate::detection_smart::detected::read_presentation_path_bytes_with_limits(
                path.as_ref(),
                limits.max_input_bytes(),
                crate::detection_smart::detected::UNIFIED_PRESENTATION_FALLBACK_MAX_INPUT_BYTES,
            )
            .map_err(Self::map_source_opc_error)?;
            return Self::from_bytes_with_limits(bytes, limits);
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
    /// - PPTX detection retains a source-backed owner and defers ordinary slide/media payloads
    /// - OLE2 and other OOXML detection return parsed owners that their loaders reuse
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
                    cached_metadata: OnceLock::from(Some(litchi_core::Metadata::default())),
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

    #[cfg(feature = "ppt")]
    fn from_native_ppt_package(package: ppt::SourceBackedPackage) -> Result<Self> {
        let cached_metadata = package
            .metadata()
            .ok()
            .map(litchi_core::Metadata::from)
            .filter(|metadata| metadata.has_data());
        let pres = package.presentation().map_err(Error::from)?;
        Ok(Self {
            inner: PresentationImpl::Ppt(pres),
            cached_metadata: OnceLock::from(cached_metadata),
        })
    }

    /// Create a presentation from bytes with an explicit PPTX/OPC resource
    /// policy. The policy is consulted only while probing an OOXML ZIP candidate.
    #[cfg(feature = "pptx")]
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: crate::pptx::ReadLimits) -> Result<Self> {
        #[cfg(feature = "odp")]
        let bytes = match crate::detection_smart::detected::detect_prepared_odp_with_limits(
            bytes, limits,
        ) {
            Ok(Ok(prepared)) => {
                return Self::from_detected(DetectedFormat::Odp(
                    prepared.into_package().into_inner(),
                ));
            },
            Ok(Err(bytes)) => bytes,
            Err(error) => return Err(Self::map_source_opc_error(error.error)),
        };

        let bytes =
            match crate::detection_smart::detected::detect_presentation_source_bytes(bytes, limits)
            {
                crate::detection_smart::detected::PresentationSourceBytesDetection::Pptx(
                    presentation,
                ) => return Self::from_source_backed_pptx(presentation),
                crate::detection_smart::detected::PresentationSourceBytesDetection::PptxError(
                    error,
                ) => return Err(map_pptx_catalog_error(error)),
                crate::detection_smart::detected::PresentationSourceBytesDetection::OpcError(
                    error,
                ) => return Err(Self::map_source_opc_error(error)),
                crate::detection_smart::detected::PresentationSourceBytesDetection::OtherOoxml(
                    format,
                ) => {
                    let _ = format;
                    return Err(Error::InvalidFormat(
                        "Detected format is not a presentation format or feature not enabled"
                            .to_owned(),
                    ));
                },
                crate::detection_smart::detected::PresentationSourceBytesDetection::DisabledOtherOoxml(
                    format,
                ) => {
                    let _ = format;
                    return Err(Error::NotOfficeFile);
                },
                crate::detection_smart::detected::PresentationSourceBytesDetection::Fallback(
                    bytes,
                ) => bytes,
            };

        let detected = crate::detection_smart::detect_format_smart_with_limits(bytes, limits)
            .ok_or(Error::NotOfficeFile)?;
        Self::from_detected(detected)
    }

    #[cfg(feature = "pptx")]
    fn from_source_backed_pptx(
        presentation: crate::pptx::SourceBackedPresentation,
    ) -> Result<Self> {
        Ok(Self {
            inner: PresentationImpl::PptxSource(presentation),
            cached_metadata: OnceLock::new(),
        })
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
                    cached_metadata: OnceLock::from(cached_metadata),
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
                    cached_metadata: OnceLock::from(cached_metadata),
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
                    cached_metadata: OnceLock::from(cached_metadata),
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
                    cached_metadata: OnceLock::from(Some(litchi_core::Metadata::default())),
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
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(crate::map_ooxml_error)?;
                let mut texts = Vec::new();
                for slide in presentation.slides() {
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
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => presentation
                .text()
                .map_err(|error| map_odp_error(error, "extract ODP text")),
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
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(crate::map_ooxml_error)?;
                Ok(presentation.slide_count())
            },
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
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => presentation
                .slide_count()
                .map_err(|error| map_odp_error(error, "get ODP slide count")),
        }
    }

    /// Return a detached point-in-time snapshot of the validated PPTX slide catalog.
    ///
    /// The returned descriptors contain each slide's zero-based position and
    /// stable `p:sldId@id`. Source-backed PPTX presentations use the retained
    /// source catalog metadata and lazy slide handles, so this method does not
    /// read slide XML or media payloads. Existing [`Self::slides`] behavior
    /// remains the eager text-bearing projection.
    ///
    /// # Errors
    ///
    /// Returns a typed source-change error or the existing facade cancellation
    /// error when a source-backed presentation is no longer current. Non-PPTX
    /// variants are unsupported because they do not expose a PPTX slide ID.
    #[cfg(feature = "pptx")]
    pub fn slide_catalog(&self) -> Result<Vec<SlideDescriptor>> {
        match &self.inner {
            PresentationImpl::Pptx(package) => {
                let presentation = package.presentation().map_err(crate::map_ooxml_error)?;
                let references = presentation
                    .slide_references()
                    .map_err(crate::map_ooxml_error)?;
                let mut descriptors = reserve_slide_catalog(references.len())?;
                for (position, reference) in references.into_iter().enumerate() {
                    descriptors.push(SlideDescriptor::new(position, reference.id()));
                }
                Ok(descriptors)
            },
            PresentationImpl::PptxSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(map_pptx_catalog_error)?;
                let mut descriptors = reserve_slide_catalog(presentation.slide_count())?;
                for slide in presentation.slides() {
                    presentation
                        .check_source()
                        .map_err(map_pptx_catalog_error)?;
                    descriptors.push(SlideDescriptor::new(slide.position(), slide.slide_id()));
                }
                presentation
                    .check_source()
                    .map_err(map_pptx_catalog_error)?;
                Ok(descriptors)
            },
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(_) => Err(Error::Unsupported(
                "PPT slide catalog requires the PPTX owner".to_owned(),
            )),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(_) => Err(Error::Unsupported(
                "slide catalog exposes PPTX slide IDs only".to_owned(),
            )),
            #[cfg(feature = "odp")]
            PresentationImpl::Odp(_) => Err(Error::Unsupported(
                "slide catalog exposes PPTX slide IDs only".to_owned(),
            )),
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(_) => Err(Error::Unsupported(
                "slide catalog exposes PPTX slide IDs only".to_owned(),
            )),
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
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => {
                use super::types::SlideData;
                presentation
                    .check_source()
                    .map_err(crate::map_ooxml_error)?;
                presentation
                    .slides()
                    .map(|slide| {
                        let (text, name) = slide.text_and_name().map_err(crate::map_ooxml_error)?;
                        Ok(Slide::Pptx(SlideData {
                            text,
                            name: Some(name),
                        }))
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
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => presentation
                .slides()
                .map(|slides| slides.into_iter().map(Slide::Odp).collect())
                .map_err(|error| map_odp_error(error, "get ODP slides")),
        }
    }

    /// Select one slide by its zero-based presentation position.
    ///
    /// The source-backed PPTX path reads only the selected slide XML and
    /// returns `None` for an out-of-range position. Other format owners retain
    /// their existing `slides()` projection as a compatibility fallback.
    pub fn slide(&self, position: usize) -> Result<Option<Slide>> {
        match &self.inner {
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(crate::map_ooxml_error)?;
                let Some(slide) = presentation.slide(position) else {
                    return Ok(None);
                };
                let (text, name) = slide.text_and_name().map_err(crate::map_ooxml_error)?;
                Ok(Some(Slide::Pptx(super::types::SlideData {
                    text,
                    name: Some(name),
                })))
            },
            #[cfg(feature = "pptx")]
            PresentationImpl::Pptx(package) => {
                let presentation = package.presentation().map_err(crate::map_ooxml_error)?;
                let Some(slide) = presentation
                    .slide(position)
                    .map_err(crate::map_ooxml_error)?
                else {
                    return Ok(None);
                };
                let text = slide.text().map_err(crate::map_ooxml_error)?;
                let name = Some(slide.name().map_err(crate::map_ooxml_error)?);
                Ok(Some(Slide::Pptx(super::types::SlideData { text, name })))
            },
            #[cfg(feature = "ppt")]
            PresentationImpl::Ppt(presentation) => {
                let Some(slide) = presentation.slide_at(position).map_err(Error::from)? else {
                    return Ok(None);
                };
                let text = slide.text().map_err(Error::from)?.to_string();
                let slide_number = slide.slide_number();
                let shape_count = slide.shape_count().unwrap_or(0);
                Ok(Some(Slide::Ppt(super::types::LegacySlideData {
                    text,
                    slide_number,
                    shape_count,
                })))
            },
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(document) => {
                let Some(slide) = document
                    .slides()
                    .map_err(|error| Error::ParseError(format!("Failed to get slides: {error}")))?
                    .get(position)
                else {
                    return Ok(None);
                };
                let title = slide.title().map(str::to_owned);
                let content = slide.text_content().join("\n");
                let text = match &title {
                    Some(title) if !content.is_empty() => format!("{title}\n\n{content}"),
                    Some(title) => title.clone(),
                    None => content,
                };
                Ok(Some(Slide::Keynote {
                    number: position + 1,
                    name: slide.name().map(str::to_owned),
                    title,
                    text,
                }))
            },
            #[cfg(feature = "odp")]
            PresentationImpl::Odp(document) => {
                let slide = document.slide(position).map_err(|error| {
                    Error::ParseError(format!("Failed to get ODP slide: {error}"))
                })?;
                Ok(slide.map(Slide::Odp))
            },
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => presentation
                .slide(position)
                .map(|slide| slide.map(Slide::Odp))
                .map_err(|error| map_odp_error(error, "get ODP slide")),
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
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => presentation
                .slide_size()
                .map(|(width, _)| Some(width))
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
            #[cfg(feature = "odp")]
            PresentationImpl::Odp(_) => Ok(None), // ODP doesn't expose slide dimensions in unified API yet
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(|error| map_odp_error(error, "check ODP source"))?;
                Ok(None)
            },
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
            #[cfg(feature = "pptx")]
            PresentationImpl::PptxSource(presentation) => presentation
                .slide_size()
                .map(|(_, height)| Some(height))
                .map_err(crate::map_ooxml_error),
            #[cfg(feature = "keynote")]
            PresentationImpl::Keynote(_) => Ok(None), // Keynote doesn't expose slide dimensions in current API
            #[cfg(feature = "odp")]
            PresentationImpl::Odp(_) => Ok(None), // ODP doesn't expose slide dimensions in unified API yet
            #[cfg(all(feature = "odp", any(unix, windows)))]
            PresentationImpl::OdpSource(presentation) => {
                presentation
                    .check_source()
                    .map_err(|error| map_odp_error(error, "check ODP source"))?;
                Ok(None)
            },
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
        #[cfg(feature = "pptx")]
        if let PresentationImpl::PptxSource(presentation) = &self.inner {
            presentation
                .check_source()
                .map_err(crate::map_ooxml_error)?;
        }

        #[cfg(all(feature = "odp", any(unix, windows)))]
        if let PresentationImpl::OdpSource(presentation) = &self.inner {
            presentation
                .check_source()
                .map_err(|error| map_odp_error(error, "check ODP source"))?;
        }

        if let Some(cached) = self.cached_metadata.get() {
            return Ok(cached.clone());
        }

        #[cfg(feature = "pptx")]
        if let PresentationImpl::PptxSource(presentation) = &self.inner {
            // Do not publish a failed read into the OnceLock: a caller can
            // retry after a transient source/version or XML failure.
            let metadata = presentation
                .properties()
                .map_err(crate::map_ooxml_error)?
                .map(litchi_core::Metadata::from)
                .filter(litchi_core::Metadata::has_data);
            let _ = self.cached_metadata.set(metadata.clone());
            return Ok(metadata);
        }

        Ok(None)
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
    #[cfg(feature = "pptx")]
    use super::PresentationImpl;
    #[cfg(feature = "doc")]
    use litchi_cfb::{OleFile, OleWriter};
    #[cfg(feature = "doc")]
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

    #[cfg(feature = "pptx")]
    #[test]
    fn native_ppt_source_owner_precedes_tiny_pptx_limit() {
        let path = test_data_path().join("ole/ppt/SampleShow.ppt");
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("consistent input limit");

        let detected = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            &path, limits,
        ) {
            Ok(Some(detected)) => detected,
            Ok(None) => panic!("native PPT source owner was not detected"),
            Err(_) => panic!("native PPT must not use the tiny PPTX limit"),
        };
        assert!(matches!(
            detected,
            crate::detection_smart::detected::PptxSourcePathDetection::Ppt(_)
        ));

        let presentation = match Presentation::open_with_limits(&path, limits) {
            Ok(presentation) => presentation,
            Err(_) => panic!("native PPT must open through its source owner"),
        };
        assert!(matches!(presentation.inner, PresentationImpl::Ppt(_)));
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

#[cfg(all(test, feature = "pptx", any(unix, windows)))]
mod source_pptx_path_tests {
    use super::super::types::PresentationImpl;
    use super::Presentation;
    use std::io::{Cursor, Write};
    use std::num::{NonZeroU64, NonZeroUsize};

    fn two_slide_package(slide1: &[u8], slide2: &[u8]) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);

        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>"#,
            )
            .unwrap();
        writer.start_file("_rels/.rels", options).unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/presentation.xml", options).unwrap();
        writer
            .write_all(
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .start_file("ppt/_rels/presentation.xml.rels", options)
            .unwrap();
        writer
            .write_all(
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.start_file("ppt/slides/slide1.xml", options).unwrap();
        writer.write_all(slide1).unwrap();
        writer.start_file("ppt/slides/slide2.xml", options).unwrap();
        writer.write_all(slide2).unwrap();
        writer.start_file("docProps/core.xml", options).unwrap();
        writer
            .write_all(
                br#"<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>source-backed</dc:title></cp:coreProperties>"#,
            )
            .unwrap();
        writer.finish().unwrap();
        output.into_inner()
    }

    fn corrupt_unselected_slide_package() -> Vec<u8> {
        two_slide_package(
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#,
            b"not PresentationML",
        )
    }

    fn late_reserved_namespace_slide_package() -> Vec<u8> {
        two_slide_package(
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:extLst xmlns:xml="urn:invalid"><p:ext uri="urn:hostile"/></p:extLst></p:sld>"#,
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#,
        )
    }

    #[test]
    fn source_slide_catalog_is_ordered_metadata_only_and_defers_bad_slide() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write PPTX fixture");

        let presentation = Presentation::open(temporary.path()).expect("source-backed PPTX");
        let before = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("filesystem PPTX must retain source owner"),
        };

        let catalog = presentation
            .slide_catalog()
            .expect("catalog must not materialize slide XML");
        assert_eq!(
            catalog
                .iter()
                .map(|slide| (slide.position(), slide.slide_id()))
                .collect::<Vec<_>>(),
            [(0, 256), (1, 257)]
        );
        let after = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("filesystem PPTX must retain source owner"),
        };
        assert_eq!(after.cold_loads, before.cold_loads);

        assert!(presentation.slide(1).is_err());
    }

    #[test]
    fn eager_slide_catalog_rejects_wrong_slide_target_like_source_catalog() {
        let input = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX input");
        std::fs::write(input.path(), corrupt_unselected_slide_package())
            .expect("write PPTX fixture");

        let mut package = crate::pptx::Package::open(input.path()).expect("eager package");
        let second_relationship_id = package
            .presentation()
            .expect("presentation root")
            .slide_references()
            .expect("slide references")
            .get(1)
            .expect("second slide reference")
            .relationship_id()
            .to_owned();
        package
            .edit_opc(|opc| {
                let presentation =
                    opc.get_part_mut(&crate::opc::PackURI::new("/ppt/presentation.xml").unwrap())?;
                presentation
                    .rels_mut()
                    .retarget(&second_relationship_id, "../docProps/core.xml".to_owned())?;
                Ok(())
            })
            .expect("retarget slide relationship");

        let output = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX output");
        package.save(output.path()).expect("save edited PPTX");
        let output_bytes = std::fs::read(output.path()).expect("read edited PPTX");
        let eager_package =
            crate::opc::OpcPackage::from_bytes(&output_bytes).expect("reopen edited OPC package");
        let presentation = Presentation::from_detected(
            crate::detection_smart::DetectedFormat::Pptx(eager_package),
        )
        .expect("eager PPTX facade");

        assert!(presentation.slide_count().is_err());
        assert!(presentation.slide_catalog().is_err());
    }

    #[test]
    fn source_slide_catalog_reports_typed_stale_source() {
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), corrupt_unselected_slide_package())
            .expect("write PPTX fixture");
        let presentation = Presentation::open(temporary.path()).expect("source-backed PPTX");

        let mut file = std::fs::OpenOptions::new()
            .append(true)
            .open(temporary.path())
            .expect("reopen PPTX for mutation");
        file.write_all(b"source mutation")
            .expect("mutate PPTX source");

        assert!(matches!(
            presentation.slide_catalog(),
            Err(litchi_core::Error::SourceChanged { .. })
        ));
    }

    #[test]
    fn path_source_routes_to_private_owner_and_defers_corrupt_slide() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write PPTX fixture");

        let direct = crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            crate::pptx::ReadLimits::default(),
        )
        .expect("source-backed private detector")
        .expect("PPTX source owner");
        assert!(matches!(
            direct,
            crate::detection_smart::detected::PptxSourcePathDetection::Pptx(_)
        ));

        let presentation = Presentation::open(temporary.path()).expect("source-backed PPTX");
        assert!(matches!(
            &presentation.inner,
            PresentationImpl::PptxSource(_)
        ));
        let before = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("path PPTX must retain source owner"),
        };
        assert_eq!(presentation.slide_count().unwrap(), 2);
        let after_count = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("path PPTX must retain source owner"),
        };
        assert_eq!(after_count.cold_loads, before.cold_loads);

        let metadata = presentation
            .metadata()
            .expect("lazy source-backed metadata")
            .expect("core properties metadata");
        assert_eq!(metadata.title.as_deref(), Some("source-backed"));
        let after_metadata = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("path PPTX must retain source owner"),
        };
        assert!(after_metadata.cold_loads > after_count.cold_loads);
        let _ = presentation.metadata().expect("cached source metadata");
        let after_cached_metadata = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("path PPTX must retain source owner"),
        };
        assert_eq!(after_cached_metadata.cold_loads, after_metadata.cold_loads);
        assert!(
            presentation
                .slide(2)
                .expect("source out-of-range query")
                .is_none()
        );

        let first = presentation
            .slide(0)
            .expect("select first source-backed slide")
            .expect("first slide exists");
        assert!(first.text().is_ok());
        let after_first = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("path PPTX must retain source owner"),
        };
        assert!(after_first.cold_loads > after_cached_metadata.cold_loads);

        assert!(presentation.slide(1).is_err());
        assert!(presentation.slides().is_err());

        let bytes_presentation = Presentation::from_bytes(bytes).expect("source-backed PPTX");
        assert!(matches!(
            &bytes_presentation.inner,
            PresentationImpl::PptxSource(_)
        ));
        assert_eq!(bytes_presentation.slide_count().unwrap(), 2);
        assert!(
            bytes_presentation
                .slide(2)
                .expect("bytes out-of-range query")
                .is_none()
        );
        assert!(bytes_presentation.slide(1).is_err());
    }

    #[test]
    fn extensionless_path_source_routes_to_private_owner() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = tempfile::NamedTempFile::new().expect("temporary extensionless PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write extensionless PPTX");

        let detected = crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            crate::pptx::ReadLimits::default(),
        )
        .expect("extensionless source-backed private detector")
        .expect("extensionless PPTX source owner");
        assert!(matches!(
            detected,
            crate::detection_smart::detected::PptxSourcePathDetection::Pptx(_)
        ));

        let presentation = Presentation::open(temporary.path()).expect("extensionless PPTX");
        assert!(matches!(
            &presentation.inner,
            PresentationImpl::PptxSource(_)
        ));
    }

    #[test]
    fn bytes_source_routes_to_private_owner_and_caches_selected_payloads() {
        let bytes = corrupt_unselected_slide_package();
        let presentation = Presentation::from_bytes(bytes).expect("source-backed PPTX bytes");
        assert!(matches!(
            &presentation.inner,
            PresentationImpl::PptxSource(_)
        ));

        let opening = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert_eq!(presentation.slide_count().unwrap(), 2);
        let after_catalog = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert_eq!(after_catalog.cold_loads, opening.cold_loads);

        assert!(
            presentation
                .metadata()
                .expect("source-backed metadata")
                .is_some()
        );
        let after_metadata = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert!(after_metadata.cold_loads > after_catalog.cold_loads);
        assert!(presentation.metadata().expect("cached metadata").is_some());
        let after_cached_metadata = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert_eq!(after_cached_metadata.cold_loads, after_metadata.cold_loads);

        let first = presentation
            .slide(0)
            .expect("first source-backed slide query")
            .expect("first slide exists");
        assert!(first.text().is_ok());
        let after_first = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert!(after_first.cold_loads > after_cached_metadata.cold_loads);

        let first_again = presentation
            .slide(0)
            .expect("cached source-backed slide query")
            .expect("first slide exists");
        assert!(first_again.text().is_ok());
        let after_cached_first = match &presentation.inner {
            PresentationImpl::PptxSource(source) => source.cache_diagnostics(),
            _ => unreachable!("bytes PPTX must retain source owner"),
        };
        assert_eq!(after_cached_first.cold_loads, after_first.cold_loads);

        // The second slide body is corrupt, but source-backed opening and
        // catalog queries do not materialize it. Selecting that slide is the
        // first operation that reports the deferred semantic error.
        assert!(presentation.slide(1).is_err());
    }

    #[test]
    fn bytes_limits_keep_source_probe_bounded() {
        let bytes = corrupt_unselected_slide_package();
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(
                u64::try_from(bytes.len() - 1).expect("fixture length fits the limit type"),
            )
            .expect("positive input limit")
            .build()
            .expect("valid input limit");

        crate::detection_smart::reset_opc_probe_count();
        let error = match Presentation::from_bytes_with_limits(bytes, limits) {
            Ok(_) => panic!("input limit must reject the PPTX"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            litchi_core::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::InputBytes
                    && limit.observed == u64::try_from(corrupt_unselected_slide_package().len()).unwrap()
                    && limit.limit == u64::try_from(corrupt_unselected_slide_package().len() - 1).unwrap()
        ));
        // The bounded private constructor is attempted once; the typed error
        // must not trigger a second eager probe.
        assert_eq!(crate::detection_smart::opc_probe_count(), 1);
    }

    #[test]
    fn source_text_rejects_reserved_namespace_prefix_consistently() {
        let bytes = late_reserved_namespace_slide_package();
        let eager = Presentation::from_detected(crate::detection_smart::DetectedFormat::Pptx(
            crate::opc::OpcPackage::from_bytes(&bytes).expect("eager OPC control"),
        ))
        .expect("eager PPTX facade control");
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write PPTX fixture");

        let source = Presentation::open(temporary.path()).expect("source-backed PPTX");
        assert!(matches!(&source.inner, PresentationImpl::PptxSource(_)));
        let bytes_presentation = Presentation::from_bytes(bytes).expect("bytes PPTX control");

        let source_error = source
            .text()
            .expect_err("reserved XML prefix must be rejected by source parsing");
        let bytes_error = bytes_presentation
            .text()
            .expect_err("reserved XML prefix must be rejected by byte parsing");
        let eager_error = eager
            .text()
            .expect_err("reserved XML prefix must be rejected by eager parsing");
        assert!(matches!(source_error, litchi_core::Error::InvalidFormat(_)));
        assert!(matches!(bytes_error, litchi_core::Error::InvalidFormat(_)));
        assert!(matches!(eager_error, litchi_core::Error::InvalidFormat(_)));
        assert!(source.slide(0).is_err());
        assert!(eager.slide(0).is_err());
    }

    #[test]
    fn cached_source_metadata_rechecks_source_revision() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write PPTX fixture");

        let presentation = Presentation::open(temporary.path()).expect("source-backed PPTX");
        assert!(presentation.metadata().unwrap().is_some());
        let mut file = std::fs::OpenOptions::new()
            .append(true)
            .open(temporary.path())
            .expect("reopen PPTX for mutation");
        file.write_all(b"source mutation")
            .expect("mutate PPTX source");

        let error = presentation
            .metadata()
            .expect_err("cached metadata must not hide a source change");
        assert!(error.to_string().contains("OPC source changed"));
    }

    #[test]
    fn cached_source_metadata_rechecks_cancellation() {
        let bytes = corrupt_unselected_slide_package();
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary PPTX path");
        std::fs::write(temporary.path(), &bytes).expect("write PPTX fixture");

        let budget = litchi_core::Budget::root(
            "litchi-presentation-metadata-test",
            litchi_core::Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        );
        let (cancellation_source, cancellation) = litchi_core::CancellationSource::pair();
        let execution_limits = litchi_core::ExecutionLimits::new(
            NonZeroUsize::new(1).expect("non-zero worker count"),
            NonZeroUsize::new(1).expect("non-zero task count"),
            NonZeroU64::new(u64::MAX).expect("non-zero in-flight bytes"),
            0,
        )
        .expect("valid execution limits");
        let context = litchi_core::ExecutionContext::new(budget, cancellation, execution_limits);
        let source =
            crate::pptx::SourceBackedPresentation::from_path_with_limits_and_execution_context(
                temporary.path(),
                crate::pptx::ReadLimits::default(),
                context,
            )
            .expect("managed source-backed PPTX");
        let presentation =
            Presentation::from_source_backed_pptx(source).expect("facade source-backed PPTX");
        assert!(presentation.metadata().unwrap().is_some());

        cancellation_source.cancel();
        assert!(matches!(
            presentation
                .slide_catalog()
                .expect_err("cancelled source catalog must fail"),
            litchi_core::Error::Other(message) if message.contains("cancel")
        ));
        let error = presentation
            .metadata()
            .expect_err("cached metadata must honor cancellation");
        assert!(error.to_string().contains("cancel"));
    }

    #[test]
    fn generic_non_zip_uses_neutral_fallback_and_public_not_office() {
        let temporary = tempfile::NamedTempFile::new().expect("temporary arbitrary input");
        std::fs::write(temporary.path(), vec![b'x'; 4096]).expect("write arbitrary input");
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("valid input limit");

        let detected = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            limits,
        ) {
            Ok(detected) => detected,
            Err(_) => panic!("generic input must use the neutral fallback ceiling"),
        };
        let Some(crate::detection_smart::detected::PptxSourcePathDetection::Bytes(bytes)) =
            detected
        else {
            panic!("generic input did not retain fallback bytes");
        };
        assert_eq!(bytes.len(), 4096);

        let error = match Presentation::open_with_limits(temporary.path(), limits) {
            Ok(_) => panic!("generic input must not become a presentation"),
            Err(error) => error,
        };
        assert!(matches!(error, litchi_core::Error::NotOfficeFile));
    }

    #[test]
    fn generic_non_zip_over_neutral_ceiling_reports_exact_limit() {
        let temporary = tempfile::NamedTempFile::new().expect("temporary sparse input");
        let neutral =
            crate::detection_smart::detected::UNIFIED_PRESENTATION_FALLBACK_MAX_INPUT_BYTES;
        let neutral_plus_one = neutral + 1;
        temporary
            .as_file()
            .set_len(neutral_plus_one)
            .expect("create sparse input");
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("consistent input limit");

        let error = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            limits,
        ) {
            Ok(_) => panic!("generic input over the neutral ceiling must fail"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            crate::detection_smart::detected::PptxSourcePathError::Opc(
                crate::opc::OpcError::ReadLimit {
                    resource: crate::opc::ReadResource::InputBytes,
                    actual,
                    maximum,
                }
            ) if actual == neutral_plus_one && maximum == neutral
        ));

        let public_error = match Presentation::open_with_limits(temporary.path(), limits) {
            Ok(_) => panic!("generic input over the neutral ceiling must fail publicly"),
            Err(error) => error,
        };
        assert!(matches!(
            public_error,
            litchi_core::Error::ResourceLimit(limit)
                if limit.resource == litchi_core::Resource::InputBytes
                    && limit.observed == neutral_plus_one
                    && limit.limit == neutral
        ));
    }

    #[test]
    fn oversized_ooxml_suffix_non_zip_keeps_eager_limit_precedence() {
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary oversized OOXML-suffixed input");
        std::fs::write(temporary.path(), vec![b'x'; 4096]).expect("write arbitrary input");
        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("valid input limit");

        let eager = crate::detection_smart::detected::detect_ooxml_path_with_limits(
            temporary.path(),
            limits,
        )
        .expect_err("eager OOXML probe must enforce the input limit first");
        assert!(matches!(
            eager,
            crate::opc::OpcError::ReadLimit {
                resource: crate::opc::ReadResource::InputBytes,
                ..
            }
        ));

        let source = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            limits,
        ) {
            Ok(_) => panic!("source-backed probe must preserve eager precedence"),
            Err(error) => error,
        };
        assert!(matches!(
            source,
            crate::detection_smart::detected::PptxSourcePathError::Opc(
                crate::opc::OpcError::ReadLimit {
                    resource: crate::opc::ReadResource::InputBytes,
                    ..
                }
            )
        ));
    }

    #[test]
    fn malformed_pptx_path_returns_typed_opc_error_without_fallback() {
        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary malformed PPTX path");
        std::fs::write(temporary.path(), b"PK\x03\x04malformed").expect("write malformed PPTX");

        let error = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            crate::pptx::ReadLimits::default(),
        ) {
            Ok(_) => panic!("malformed OOXML-suffixed ZIP must remain typed"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            crate::detection_smart::detected::PptxSourcePathError::Opc(
                crate::opc::OpcError::ZipError(_)
            )
        ));
    }

    #[test]
    fn path_part_limit_returns_typed_opc_error_with_exact_bounds() {
        use std::io::Write;

        let mut output = std::io::Cursor::new(Vec::new());
        let mut writer = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        writer.start_file("[Content_Types].xml", options).unwrap();
        writer
            .write_all(
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="txt" ContentType="text/plain"/></Types>"#,
            )
            .unwrap();
        writer.start_file("one.txt", options).unwrap();
        writer.write_all(b"one").unwrap();
        writer.start_file("two.txt", options).unwrap();
        writer.write_all(b"two").unwrap();
        writer.finish().unwrap();

        let temporary = tempfile::Builder::new()
            .suffix(".pptx")
            .tempfile()
            .expect("temporary part-limited PPTX path");
        std::fs::write(temporary.path(), output.into_inner()).expect("write part-limited ZIP");
        let limits = crate::pptx::ReadLimits::builder()
            .max_parts(1)
            .expect("test part limit")
            .build()
            .expect("consistent test limits");

        let error = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            temporary.path(),
            limits,
        ) {
            Ok(_) => panic!("part limit must stop the source-backed path probe"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            crate::detection_smart::detected::PptxSourcePathError::Opc(
                crate::opc::OpcError::ReadLimit {
                    resource: crate::opc::ReadResource::Parts,
                    actual: 2,
                    maximum: 1,
                }
            )
        ));
    }

    #[cfg(feature = "odp")]
    #[test]
    fn source_probe_returns_odp_owner_and_same_source_unknown_zip_fallback() {
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
        let odp = writer.finish_to_bytes().unwrap();
        let odp_path = tempfile::Builder::new()
            .suffix(".odp")
            .tempfile()
            .expect("temporary ODP path");
        std::fs::write(odp_path.path(), odp).expect("write ODP");
        let odp_detected = crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            odp_path.path(),
            crate::pptx::ReadLimits::default(),
        )
        .expect("ODP source probe");
        assert!(matches!(
            odp_detected,
            Some(crate::detection_smart::detected::PptxSourcePathDetection::Odp(_))
        ));
        let odp_presentation = Presentation::open(odp_path.path()).expect("open ODP");
        assert!(matches!(
            &odp_presentation.inner,
            PresentationImpl::OdpSource(_)
        ));

        let mut output = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(&mut output);
        zip.start_file(
            "plain.txt",
            zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored),
        )
        .unwrap();
        zip.write_all(b"not an Office package").unwrap();
        zip.finish().unwrap();
        let unknown_bytes = output.into_inner();
        let unknown_path = tempfile::Builder::new()
            .suffix(".zip")
            .tempfile()
            .expect("temporary unknown ZIP path");
        std::fs::write(unknown_path.path(), &unknown_bytes).expect("write unknown ZIP");
        let detected = crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            unknown_path.path(),
            crate::pptx::ReadLimits::default(),
        )
        .expect("extensionless ZIP fallback probe")
        .expect("same-source fallback bytes");
        let crate::detection_smart::detected::PptxSourcePathDetection::Bytes(retained) = detected
        else {
            panic!("unknown ZIP did not retain same-source fallback bytes");
        };
        assert_eq!(retained, unknown_bytes);

        let limits = crate::pptx::ReadLimits::builder()
            .max_input_bytes(1)
            .expect("positive input limit")
            .build()
            .expect("consistent input limit");
        let error = match crate::detection_smart::detected::detect_pptx_source_path_with_limits(
            unknown_path.path(),
            limits,
        ) {
            Ok(_) => panic!("unknown ZIP must remain caller-limited"),
            Err(error) => error,
        };
        assert!(matches!(
            error,
            crate::detection_smart::detected::PptxSourcePathError::Opc(
                crate::opc::OpcError::ReadLimit {
                    resource: crate::opc::ReadResource::InputBytes,
                    actual,
                    maximum: 1,
                }
            ) if actual == u64::try_from(unknown_bytes.len()).unwrap()
        ));
        let error = match Presentation::open(unknown_path.path()) {
            Ok(_) => panic!("unknown ZIP must remain unsupported"),
            Err(error) => error,
        };
        assert!(matches!(error, litchi_core::Error::NotOfficeFile));
    }
}
