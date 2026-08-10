//! Immutable PresentationML reads backed by a caller-provided positional source.
//!
//! This facade keeps ordinary slide payloads deferred. Opening validates the
//! OPC catalog and mandatory presentation root, then resolves only slide
//! metadata. A slide body is loaded when a selected [`SourceSlide`] is read.

use std::sync::{Arc, OnceLock};

use litchi_core::ReadAt;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, PackURI, Part, PartData, PartView, ReadLimits, SourceBackedPackage};

use crate::parts::{PresentationPart, SlidePart, SlideReference};
use crate::{Error, Result};

struct SourceSlideData {
    position: usize,
    part_uri: PackURI,
    part: OnceLock<BlobPart>,
}

struct SourceInner {
    package: SourceBackedPackage,
    // Retain the pinned mandatory root without relying on the OPC payload cache.
    _presentation: BlobPart,
    slides: Box<[Arc<SourceSlideData>]>,
}

/// Read-only PPTX catalog and selected-slide access over a positional source.
///
/// Opening validates the OPC catalog, package relationships, presentation
/// part, and presentation-to-slide graph. Slide payloads remain deferred until
/// [`SourceSlide::text`] selects one. The type has no edit or output APIs.
#[derive(Clone)]
pub struct SourceBackedPresentation {
    inner: Arc<SourceInner>,
}

/// A lifetime-free read-only slide handle from [`SourceBackedPresentation`].
///
/// Creating or listing handles does not read slide XML.
#[derive(Clone)]
pub struct SourceSlide {
    owner: Arc<SourceInner>,
    data: Arc<SourceSlideData>,
}

impl SourceBackedPresentation {
    /// Open an ordinary PPTX package from a caller-provided positional source.
    ///
    /// # Errors
    ///
    /// Returns an error when the source, OPC catalog, presentation root, or
    /// presentation-to-slide graph is malformed.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default())
    }

    /// Open from a positional source with explicit OPC resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the source exceeds `limits`, changes while being
    /// read, or does not contain a coherent PresentationML catalog.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Build the read-only PPTX facade from a validated deferred OPC package.
    ///
    /// Only the mandatory presentation payload is read by this constructor.
    ///
    /// # Errors
    ///
    /// Returns an error when the main part or its ordered slide graph is not a
    /// coherent PresentationML presentation.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        let (presentation, slides) = {
            let view = package.main_document_part()?;
            let presentation = owned_part(&view, view.data()?)?;
            let references = PresentationPart::from_part(&presentation)?.slide_references()?;
            let slides = validate_slide_graph(&package, &view, &references)?;
            (presentation, slides)
        };

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                _presentation: presentation,
                slides,
            }),
        })
    }

    /// Number of logical slides in presentation order.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.inner.slides.len()
    }

    /// Iterate lightweight slide handles without reading slide payloads.
    #[must_use]
    pub fn slides(&self) -> impl ExactSizeIterator<Item = SourceSlide> + DoubleEndedIterator + '_ {
        self.inner.slides.iter().cloned().map(|data| SourceSlide {
            owner: Arc::clone(&self.inner),
            data,
        })
    }

    /// Select one slide by checked zero-based presentation position.
    #[must_use]
    pub fn slide(&self, position: usize) -> Option<SourceSlide> {
        let data = self.inner.slides.get(position)?.clone();
        Some(SourceSlide {
            owner: Arc::clone(&self.inner),
            data,
        })
    }
}

impl SourceSlide {
    /// Checked zero-based position in the presentation catalog.
    #[must_use]
    pub fn position(&self) -> usize {
        self.data.position
    }

    /// Flatten DrawingML text runs from this selected slide in source order.
    ///
    /// The slide payload is loaded and retained on first use. No other slide
    /// payload is read.
    ///
    /// # Errors
    ///
    /// Returns an error if the source changed, the selected slide exceeds the
    /// retained OPC limits, or its PresentationML is malformed.
    pub fn text(&self) -> Result<String> {
        SlidePart::from_part(self.part()?)?.text()
    }

    fn part(&self) -> Result<&BlobPart> {
        // The metadata lookup keeps source-version checks active even after a
        // selected slide payload has entered the local cache.
        let view = self.owner.package.part(&self.data.part_uri)?;
        if let Some(part) = self.data.part.get() {
            return Ok(part);
        }

        let part = owned_part(&view, view.data()?)?;
        let _publish_result = self.data.part.set(part);
        self.data.part.get().ok_or_else(|| {
            Error::Invalid("source-backed slide cache did not publish a value".to_string())
        })
    }
}

fn validate_slide_graph(
    package: &SourceBackedPackage,
    presentation: &PartView<'_>,
    references: &[SlideReference],
) -> Result<Box<[Arc<SourceSlideData>]>> {
    let mut slides = Vec::new();
    slides
        .try_reserve_exact(references.len())
        .map_err(|source| Error::Allocation {
            resource: "source-backed presentation slide graph",
            source,
        })?;
    for (position, reference) in references.iter().enumerate() {
        let relationship = presentation
            .rels()
            .get(reference.relationship_id())
            .ok_or_else(|| {
                Error::Relationship(format!(
                    "presentation slide reference is missing relationship '{}'",
                    reference.relationship_id()
                ))
            })?;
        if relationship.is_external() {
            return Err(Error::Relationship(format!(
                "slide relationship '{}' must be internal",
                reference.relationship_id()
            )));
        }
        if !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide") {
            return Err(Error::Relationship(format!(
                "relationship '{}' is not a slide relationship",
                reference.relationship_id()
            )));
        }
        let part_uri = relationship.target_partname()?;
        let slide = package.part(&part_uri)?;
        if slide.content_type() != ct::PML_SLIDE {
            return Err(Error::ContentType {
                expected: ct::PML_SLIDE.to_string(),
                actual: slide.content_type().to_string(),
            });
        }
        slides.push(Arc::new(SourceSlideData {
            position,
            part_uri,
            part: OnceLock::new(),
        }));
    }
    Ok(slides.into_boxed_slice())
}

fn owned_part(view: &PartView<'_>, data: PartData) -> Result<BlobPart> {
    let mut part = BlobPart::new_shared(
        view.partname().clone(),
        view.content_type().to_string(),
        data.into_arc(),
    );
    for relationship in view.rels().iter() {
        part.rels_mut().try_add_relationship(
            relationship.reltype().to_string(),
            relationship.target_ref().to_string(),
            relationship.r_id().to_string(),
            relationship.target_mode(),
        )?;
    }
    Ok(part)
}

#[cfg(test)]
mod tests {
    use std::io;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

    use litchi_core::{ReadAt, SourceVersion};
    use litchi_opc::{ReadLimits, ReadResource};
    use soapberry_zip::office::StreamingArchiveWriter;

    use super::SourceBackedPresentation;
    use crate::Error;

    const SECOND_MARKER: &[u8] = b"source-backed-unrequested-second-slide";

    struct CountingSource {
        bytes: Vec<u8>,
        marker_offset: usize,
        second_payload_reads: AtomicUsize,
        revision: AtomicU64,
    }

    impl CountingSource {
        fn new(bytes: Vec<u8>) -> Self {
            let marker_offset = bytes
                .windows(SECOND_MARKER.len())
                .position(|window| window == SECOND_MARKER)
                .expect("second slide marker is stored in archive");
            Self {
                bytes,
                marker_offset,
                second_payload_reads: AtomicUsize::new(0),
                revision: AtomicU64::new(0),
            }
        }

        fn changed(&self) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
    }

    impl ReadAt for CountingSource {
        fn len(&self) -> io::Result<u64> {
            Ok(self.bytes.len() as u64)
        }

        fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
            let offset = usize::try_from(offset)
                .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
            if offset >= self.bytes.len() {
                return Ok(0);
            }
            let count = output.len().min(self.bytes.len() - offset);
            let end = offset + count;
            if offset < self.marker_offset + SECOND_MARKER.len() && self.marker_offset < end {
                self.second_payload_reads.fetch_add(1, Ordering::SeqCst);
            }
            output[..count].copy_from_slice(&self.bytes[offset..end]);
            Ok(count)
        }

        fn version(&self) -> io::Result<SourceVersion> {
            Ok(SourceVersion::new(92, self.revision.load(Ordering::SeqCst)))
        }
    }

    fn source_backed_pptx() -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/presentation.xml",
                br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/_rels/presentation.xml.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>"#,
            )
            .unwrap();
        writer
            .write_stored(
                "ppt/slides/slide1.xml",
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><a:t>First slide</a:t></p:spTree></p:cSld></p:sld>"#,
            )
            .unwrap();
        let padding = "x".repeat(128 * 1024);
        let second = format!(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><!--{marker}{padding}--><p:cSld><p:spTree><a:t>Second slide</a:t></p:spTree></p:cSld></p:sld>"#,
            marker = std::str::from_utf8(SECOND_MARKER).unwrap(),
        );
        writer
            .write_stored("ppt/slides/slide2.xml", second.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn catalog_and_selected_text_leave_unselected_slides_unread() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();

        assert_eq!(presentation.slide_count(), 2);
        assert_eq!(
            presentation
                .slides()
                .map(|slide| slide.position())
                .collect::<Vec<_>>(),
            [0, 1]
        );
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);

        let first = presentation.slide(0).unwrap();
        assert_eq!(first.text().unwrap(), "First slide");
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);

        let second = presentation.slide(1).unwrap();
        assert_eq!(second.text().unwrap(), "Second slide");
        assert!(source.second_payload_reads.load(Ordering::SeqCst) > 0);
    }

    #[test]
    fn source_changes_are_returned_as_typed_opc_errors() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
        let slide = presentation.slide(0).unwrap();
        source.changed();

        assert!(matches!(
            slide.text(),
            Err(Error::Opc(litchi_opc::OpcError::SourceChanged { .. }))
        ));
    }

    #[test]
    fn opening_retains_caller_part_limits_without_reading_slide_payloads() {
        let source = Arc::new(CountingSource::new(source_backed_pptx()));
        let limits = ReadLimits::builder()
            .max_part_bytes(1024)
            .unwrap()
            .build()
            .unwrap();

        assert!(matches!(
            SourceBackedPresentation::from_read_at_with_limits(source.clone(), limits),
            Err(Error::Opc(litchi_opc::OpcError::ReadLimit {
                resource: ReadResource::PartBytes,
                maximum: 1024,
                ..
            }))
        ));
        assert_eq!(source.second_payload_reads.load(Ordering::SeqCst), 0);
    }
}
