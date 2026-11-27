//! Detached, source-backed presentation sections.
//!
//! Section XML is parsed through the existing metadata codec after the
//! source owner has applied its physical and markup-compatibility boundary.
//! The returned values contain semantic slide positions only; package IDs and
//! inert extension bytes require the explicit [`raw`] module.
//!
//! ```
//! use std::io::Cursor;
//! use litchi_pptx::{Package, SourceBackedPresentation};
//!
//! let mut package = Package::new()?;
//! let source = SourceBackedPresentation::from_reader(Cursor::new(package.to_bytes()?))?;
//! let sections = source.sections()?;
//! drop(source);
//! assert!(sections.is_empty());
//! # Ok::<(), litchi_pptx::Error>(())
//! ```

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_opc::Part;

use crate::error::{Error, Result};
use crate::presentation_properties::metadata::sections::List as MetadataList;

use super::SourceInner;

/// An immutable, detached section projection.
#[derive(Clone)]
pub struct List {
    data: Arc<ListData>,
}

struct ListData {
    entries: Box<[Section]>,
}

impl fmt::Debug for List {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("List")
            .field("len", &self.data.entries.len())
            .finish_non_exhaustive()
    }
}

impl List {
    /// Sections in authored presentation order.
    #[must_use]
    pub fn entries(&self) -> &[Section] {
        &self.data.entries
    }

    /// Return the section at a checked zero-based position.
    #[must_use]
    pub fn section(&self, position: Position) -> Option<&Section> {
        self.data.entries.get(position.get())
    }

    /// Number of sections in this projection.
    #[must_use]
    pub fn len(&self) -> usize {
        self.data.entries.len()
    }

    /// Whether the presentation contains no sections.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.entries.is_empty()
    }
}

/// One immutable semantic presentation section.
#[derive(Clone)]
pub struct Section {
    data: Arc<SectionData>,
}

struct SectionData {
    position: Position,
    name: Option<String>,
    slides: Box<[Position]>,
    native_slide_ids: Box<[u32]>,
    id: Option<String>,
    extension_xml: Option<Vec<u8>>,
}

impl fmt::Debug for Section {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        // Names, GUIDs, and opaque XML are authored document data. Keep the
        // ordinary diagnostic shape useful without echoing any of them.
        formatter
            .debug_struct("Section")
            .field("position", &self.data.position)
            .field("slide_count", &self.data.slides.len())
            .finish_non_exhaustive()
    }
}

impl Section {
    /// The optional authored display name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.data.name.as_deref()
    }

    /// Slides in this section's authored order.
    #[must_use]
    pub fn slides(&self) -> &[Position] {
        &self.data.slides
    }

    /// This section's checked zero-based position in the section list.
    #[must_use]
    pub fn position(&self) -> Position {
        self.data.position
    }
}

/// Explicit raw escape hatches for native section identity and inert XML.
pub mod raw {
    use super::Section;

    /// Return the producer-supplied section GUID, if present.
    #[must_use]
    pub fn id(section: &Section) -> Option<&str> {
        section.data.id.as_deref()
    }

    /// Return the producer-supplied `p14:sldId` values in authored order.
    #[must_use]
    pub fn native_slide_ids(section: &Section) -> &[u32] {
        &section.data.native_slide_ids
    }

    /// Return the inert serialized section extension list, if present.
    #[must_use]
    pub fn extension_xml(section: &Section) -> Option<&[u8]> {
        section.data.extension_xml.as_deref()
    }
}

impl super::SourceBackedPresentation {
    /// Read and detach the presentation's semantic sections.
    ///
    /// Section XML is loaded only when this method is called. The returned
    /// list remains readable after the source owner is dropped, while a
    /// source revision or execution cancellation observed during acquisition
    /// refuses publication.
    ///
    /// # Errors
    /// Returns typed source, execution, allocation, limit or invalid-section errors.
    pub fn sections(&self) -> Result<List> {
        super::related::with_fences(&self.inner, || read_sections(&self.inner))
    }
}

fn read_sections(owner: &SourceInner) -> Result<List> {
    let part = super::related::load_part(
        owner,
        owner._presentation.source_partname(),
        owner._presentation.content_type(),
        usize::try_from(owner.limits.max_part_bytes()).unwrap_or(usize::MAX),
    )?;
    let metadata = MetadataList::from_processed(
        part.blob(),
        crate::presentation_properties::metadata::sections::Limits::from_source_limits(
            owner.limits,
        ),
    )
    .map_err(|error| match error {
        // The legacy codec includes a few parser token names in diagnostics.
        // Keep this canonical source surface content-free while preserving
        // typed allocation/OPC failures from the owner boundary.
        Error::Invalid(_) => Error::Invalid("source presentation sections XML is invalid".into()),
        Error::Xml(_) => Error::Xml("source presentation sections XML is malformed".into()),
        error => error,
    })?;
    project_sections(owner, metadata)
}

fn project_sections(owner: &SourceInner, metadata: MetadataList) -> Result<List> {
    let mut positions = HashMap::new();
    positions
        .try_reserve(owner.slides.len())
        .map_err(|source| Error::Allocation {
            resource: "source section slide index",
            source,
        })?;
    for (position, slide) in owner.slides.iter().enumerate() {
        if positions
            .insert(slide.slide_id, Position::new(position))
            .is_some()
        {
            return Err(Error::Invalid(
                "source presentation has duplicate native slide IDs".into(),
            ));
        }
    }

    let mut entries = Vec::new();
    entries
        .try_reserve_exact(metadata.sections.len())
        .map_err(|source| Error::Allocation {
            resource: "source presentation sections",
            source,
        })?;

    for (section_position, section) in metadata.sections.into_iter().enumerate() {
        let mut slides = Vec::new();
        slides
            .try_reserve_exact(section.slide_ids.len())
            .map_err(|source| Error::Allocation {
                resource: "source section slide positions",
                source,
            })?;
        for slide_id in &section.slide_ids {
            let Some(position) = positions.get(slide_id).copied() else {
                return Err(Error::Invalid(
                    "source section references an unknown presentation slide".into(),
                ));
            };
            slides.push(position);
        }
        entries.push(Section {
            data: Arc::new(SectionData {
                position: Position::new(section_position),
                name: section.name,
                slides: slides.into_boxed_slice(),
                native_slide_ids: section.slide_ids.into_boxed_slice(),
                id: section.id,
                extension_xml: section.extension_xml,
            }),
        });
    }

    Ok(List {
        data: Arc::new(ListData {
            entries: entries.into_boxed_slice(),
        }),
    })
}
