//! Native Keynote package ingress and semantic decoding.
//!
//! This adapter owns the `.key` package boundary while preserving every raw
//! package member in its original byte stream. The archive, Snappy, detection,
//! and protobuf layers remain in their focused IWA infrastructure crates.

mod edit;
mod limits;
mod show_settings;
mod slide_order;

use std::fmt;
use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;
use std::sync::Arc;
#[cfg(test)]
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_iwa_archive::{ComponentCatalog, Limits as ArchiveLimits, SourceCatalog};
use litchi_iwa_common::{
    WireLimits,
    wire::{WireDescent, WireFieldView, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_detect::Format;
#[cfg(feature = "internal-iwork-source")]
use litchi_iwa_detect::PreparedSource;
use litchi_iwa_protos::{keynote_document_codec, keynote_show_codec, kn, tswp};
use litchi_iwa_text::storage::Storage;
use litchi_iwa_text_wire::{
    DEFAULT_MAX_FIELDS as DEFAULT_MAX_TEXT_FIELDS,
    DEFAULT_MAX_WIRE_FRAGMENTS as DEFAULT_MAX_TEXT_FRAGMENTS, Error as TextWireError,
    Limits as TextWireLimits,
};
use once_cell::sync::OnceCell;
use prost::Message;
use thiserror::Error;

use crate::{
    AnimationType, Build, Document, Effect, Mode, Seconds, Settings, Show, Size, Slide, Transition,
};

pub use edit::{Commit, Diagnostics, Edit, EditError, Patch};
pub use limits::{
    MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES,
    ReadOptions, SemanticLimitKind, SemanticLimits, SemanticLimitsError,
};
pub use show_settings::{
    ShowSettingsCommit, ShowSettingsDiagnostics, ShowSettingsEdit, ShowSettingsError,
    ShowSettingsLimitKind, ShowSettingsPatch,
};
pub use slide_order::{
    SlideOrderCommit, SlideOrderDiagnostics, SlideOrderEdit, SlideOrderError, SlideOrderLimitKind,
    SlideOrderPatch,
};

/// Checked physical resource limits for Keynote package ingress.
pub use litchi_iwa_archive::Limits;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const BUILD_MESSAGE_TYPE: u32 = 8;
const NOTE_MESSAGE_TYPE: u32 = 15;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;

/// A result returned by a native Keynote package operation.
type ReadResult<T> = Result<T, ReadError>;

/// Content-free semantic location associated with a Keynote read failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticPath {
    /// Whole-package ingress or indexing.
    Package,
    /// The presentation show root.
    Show,
    /// The optional show title.
    ShowTitle,
    /// One slide at a semantic zero-based position.
    Slide { index: usize },
    /// One slide's semantic navigation name.
    SlideName { index: usize },
    /// One slide's title placeholder.
    SlideTitle { index: usize },
    /// One slide's body placeholder.
    SlideBody { index: usize },
    /// One non-placeholder drawable in slide source order.
    SlideDrawable { slide: usize, index: usize },
    /// One slide's speaker notes.
    SlideNotes { index: usize },
    /// One build in slide source order.
    SlideBuild { slide: usize, index: usize },
    /// One slide's transition.
    SlideTransition { index: usize },
}

impl fmt::Display for SemanticPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Show => formatter.write_str("show"),
            Self::ShowTitle => formatter.write_str("show title"),
            Self::Slide { index } => write!(formatter, "slide {index}"),
            Self::SlideName { index } => write!(formatter, "slide {index} name"),
            Self::SlideTitle { index } => write!(formatter, "slide {index} title"),
            Self::SlideBody { index } => write!(formatter, "slide {index} body"),
            Self::SlideDrawable { slide, index } => {
                write!(formatter, "slide {slide} drawable {index}")
            },
            Self::SlideNotes { index } => write!(formatter, "slide {index} notes"),
            Self::SlideBuild { slide, index } => {
                write!(formatter, "slide {slide} build {index}")
            },
            Self::SlideTransition { index } => write!(formatter, "slide {index} transition"),
        }
    }
}

/// A bounded native payload resource reported without leaking wire-layer types.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum PayloadLimitKind {
    /// Encoded or rewritten payload bytes.
    Bytes,
    /// Parsed fields or field-like records.
    Fields,
    /// Nested message traversal depth.
    Nesting,
    /// Aggregate traversal or rewrite work.
    Work,
}

impl fmt::Display for PayloadLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Bytes => "bytes",
            Self::Fields => "fields",
            Self::Nesting => "nesting depth",
            Self::Work => "work",
        })
    }
}

/// An error raised while reading or decoding a Keynote package.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum ReadError {
    /// Reading a package from the filesystem failed.
    #[error("could not read Keynote package: {0}")]
    Io(#[from] std::io::Error),
    /// The physical iWork package boundary rejected the input.
    #[error(transparent)]
    Archive(#[from] litchi_iwa_archive::Error),
    /// iWork format detection rejected the input.
    #[error(transparent)]
    Detection(#[from] litchi_iwa_detect::Error),
    /// The package is valid iWork data but is not a Keynote presentation.
    #[error("iWork package is not a Keynote presentation")]
    NotKeynote,
    /// The package does not contain the Keynote structure required by this reader.
    #[error("invalid Keynote package: {0}")]
    InvalidFormat(String),
    /// A native Keynote payload could not be translated into its semantic value.
    #[error("could not decode Keynote content: {0}")]
    Decode(String),
    /// A package-wide semantic resource ceiling was exceeded.
    #[error(
        "Keynote semantic {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    SemanticLimit {
        /// Resource category that exceeded its ceiling.
        kind: SemanticLimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
        /// Content-free semantic location where the limit was encountered.
        path: SemanticPath,
    },
    /// A bounded native payload preflight exceeded its finite profile.
    #[error(
        "Keynote payload {kind} limit exceeded at {path}: observed {observed}, maximum {maximum}"
    )]
    PayloadLimit {
        /// Runtime-neutral payload resource category.
        kind: PayloadLimitKind,
        /// Observed or requested amount.
        observed: usize,
        /// Configured maximum.
        maximum: usize,
        /// Content-free semantic location where the limit was encountered.
        path: SemanticPath,
    },
    /// A destination allocation failed before semantic state was published.
    #[error("could not allocate {amount} units for {resource}")]
    Allocation {
        /// Stable semantic allocation category.
        resource: &'static str,
        /// Elements or bytes requested.
        amount: usize,
    },
    /// A native text-storage payload failed strict bounded projection.
    #[error("invalid Keynote text storage at {path}: {reason}")]
    TextStorage {
        /// Stable, content-free failure category.
        reason: TextStorageFailure,
        /// Content-free semantic location of the referenced storage.
        path: SemanticPath,
    },
    /// The package properties plist could not be read.
    #[error("could not read Keynote package properties: {0}")]
    Metadata(#[from] plist::Error),
}

/// Why a recognized native Keynote text-storage payload was rejected.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextStorageFailure {
    /// A text fragment is not valid UTF-8.
    InvalidUtf8,
    /// The native text field has the wrong protobuf wire type.
    WrongWireType,
    /// The bounded Buffa projection disagreed with validated raw wire data.
    Projection,
    /// Allocation of the semantic text value failed.
    Allocation,
    /// The projected text/range relation is invalid.
    InvalidRanges,
    /// The storage wire representation is otherwise malformed.
    MalformedWire,
}

impl fmt::Display for TextStorageFailure {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InvalidUtf8 => "invalid UTF-8",
            Self::WrongWireType => "wrong text wire type",
            Self::Projection => "lazy projection mismatch",
            Self::Allocation => "semantic allocation failed",
            Self::InvalidRanges => "invalid semantic text ranges",
            Self::MalformedWire => "malformed protobuf wire data",
        })
    }
}

/// Cheaply cloneable parsed Keynote package with a lazy semantic snapshot.
///
/// The original package bytes remain available through [`Self::source_bytes`],
/// so unsupported IWA members and unmodeled protobuf fields are retained even
/// when callers inspect only the semantic presentation values.
#[derive(Clone)]
pub struct Package {
    state: Arc<State>,
}

impl fmt::Debug for Package {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Package").finish_non_exhaustive()
    }
}

#[derive(Debug)]
struct State {
    source: PhysicalSource,
    options: ReadOptions,
    object_index: Box<[ObjectLocator]>,
    total_objects: usize,
    semantic: OnceCell<Document>,
    #[cfg(test)]
    semantic_decode_attempts: AtomicUsize,
    #[cfg(all(test, feature = "internal-iwork-source"))]
    source_classification_attempts: usize,
}

#[derive(Debug)]
enum PhysicalSource {
    Package(SourceCatalog),
    #[cfg(feature = "internal-iwork-source")]
    Semantic(Arc<ComponentCatalog>),
}

impl PhysicalSource {
    fn components(&self) -> &ComponentCatalog {
        match self {
            Self::Package(source) => source.components(),
            #[cfg(feature = "internal-iwork-source")]
            Self::Semantic(components) => components,
        }
    }

    fn package_source(&self) -> &SourceCatalog {
        match self {
            Self::Package(source) => source,
            #[cfg(feature = "internal-iwork-source")]
            Self::Semantic(_) => {
                panic!("semantic-only Keynote decoder has no physical package source")
            },
        }
    }

    fn package(&self) -> &litchi_iwa_archive::package::Catalog {
        self.package_source().package()
    }

    fn source_bytes(&self) -> &[u8] {
        self.package_source().source_bytes()
    }

    fn shared_source(&self) -> Arc<[u8]> {
        self.package_source().shared_source()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ObjectLocator {
    identifier: u64,
    component: usize,
    object: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SlideRecord {
    node_identifier: u64,
    slide_identifier: u64,
    is_skipped: bool,
}

/// Deterministic measurements for one Keynote package snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Stats {
    /// Number of native IWA objects retained by the parsed package.
    pub total_objects: usize,
    /// Number of semantic slides resolved from the Keynote show tree.
    pub slide_count: usize,
}

impl Package {
    /// Open a Keynote package from a filesystem path with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the file cannot be read, is not a bounded valid
    /// iWork package, or does not contain a Keynote document root.
    pub fn open(path: impl AsRef<Path>) -> ReadResult<Self> {
        Self::open_with_options(path, ReadOptions::default())
    }

    /// Open a Keynote package from a filesystem path with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`] while enforcing `limits`
    /// before materializing the complete source and at the physical ZIP and
    /// IWA boundaries.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> ReadResult<Self> {
        Self::open_with_options(path, ReadOptions::new(limits, SemanticLimits::default()))
    }

    /// Open a Keynote package with independent physical and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`]. Physical and object-index
    /// limits are enforced before this returns; slide, reference, and text
    /// limits are enforced lazily on first semantic access or [`Self::validate`].
    pub fn open_with_options(path: impl AsRef<Path>, options: ReadOptions) -> ReadResult<Self> {
        let source = read_source(path.as_ref(), options.archive())?;
        Self::from_source_with_options(source, options)
    }

    /// Parse a Keynote package from complete ZIP bytes with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is not a bounded valid Keynote package.
    pub fn from_bytes(bytes: &[u8]) -> ReadResult<Self> {
        Self::from_bytes_with_options(bytes, ReadOptions::default())
    }

    /// Parse a Keynote package from complete ZIP bytes with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP, Snappy/IWA components, document root,
    /// or requested resource profile is invalid.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> ReadResult<Self> {
        Self::from_bytes_with_options(bytes, ReadOptions::new(limits, SemanticLimits::default()))
    }

    /// Parse complete Keynote ZIP bytes with independent physical and semantic limits.
    ///
    /// # Errors
    ///
    /// Returns an error when physical ingress, format detection, or the object
    /// index exceeds its profile. Remaining semantic limits are enforced lazily
    /// on first semantic access or [`Self::validate`].
    pub fn from_bytes_with_options(bytes: &[u8], options: ReadOptions) -> ReadResult<Self> {
        let limits = options.archive();
        check_input_size(
            u64::try_from(bytes.len()).map_err(|_error| {
                ReadError::InvalidFormat("Keynote input length does not fit u64".to_owned())
            })?,
            limits,
        )?;
        let source = copy_source(bytes)?;
        Self::from_source_with_options(source, options)
    }

    /// Parse Keynote package bytes supplied through an archive-oriented API.
    ///
    /// This is equivalent to [`Self::from_bytes`].
    ///
    /// # Errors
    ///
    /// Returns an error when `bytes` is not a bounded valid Keynote package.
    pub fn from_archive_bytes(bytes: &[u8]) -> ReadResult<Self> {
        Self::from_bytes(bytes)
    }

    fn from_source_with_options(source: Arc<[u8]>, options: ReadOptions) -> ReadResult<Self> {
        let limits = options.archive();
        let source_catalog = SourceCatalog::from_shared_bytes_with_limits(source, limits)?;
        Self::from_source_catalog(source_catalog, options.semantic())
    }

    /// Consume one source prepared by the focused iWork coordinator.
    ///
    /// This explicitly unstable handoff preserves the original immutable
    /// package allocation and its validated physical profile. Only the
    /// semantic profile is selected at this stage.
    ///
    /// # Errors
    ///
    /// Returns [`ReadError`] when the source belongs to another application or
    /// the Keynote root/index cannot be validated.
    #[cfg(feature = "internal-iwork-source")]
    #[doc(hidden)]
    pub fn __from_prepared_source(
        source: PreparedSource,
        semantic: SemanticLimits,
    ) -> ReadResult<Self> {
        if source.format() != Format::Keynote {
            return Err(ReadError::NotKeynote);
        }
        let source_catalog = source.__into_source_catalog().ok_or_else(|| {
            ReadError::InvalidFormat(
                "directory-backed Keynote sources support semantic projection only".to_owned(),
            )
        })?;
        Self::from_classified_source_catalog(source_catalog, semantic, 0)
    }

    fn from_source_catalog(
        source_catalog: SourceCatalog,
        semantic: SemanticLimits,
    ) -> ReadResult<Self> {
        match litchi_iwa_detect::component_catalog(source_catalog.components())? {
            Some(Format::Keynote) => {},
            Some(_) => return Err(ReadError::NotKeynote),
            None => {
                return Err(ReadError::InvalidFormat(
                    "package has no recognized iWork application root".to_owned(),
                ));
            },
        }

        Self::from_classified_source_catalog(source_catalog, semantic, 1)
    }

    fn from_classified_source_catalog(
        source_catalog: SourceCatalog,
        semantic: SemanticLimits,
        source_classification_attempts: usize,
    ) -> ReadResult<Self> {
        debug_assert!(source_classification_attempts <= 1);
        let (object_index, total_objects) =
            build_object_index(source_catalog.components(), semantic.max_objects())?;
        let options = ReadOptions::new(source_catalog.limits(), semantic);

        let package = Self {
            state: Arc::new(State {
                source: PhysicalSource::Package(source_catalog),
                options,
                object_index,
                total_objects,
                semantic: OnceCell::new(),
                #[cfg(test)]
                semantic_decode_attempts: AtomicUsize::new(0),
                #[cfg(all(test, feature = "internal-iwork-source"))]
                source_classification_attempts,
            }),
        };
        package.root_show_identifier()?;
        Ok(package)
    }

    #[cfg(feature = "internal-iwork-source")]
    fn from_classified_components(
        components: Arc<ComponentCatalog>,
        archive: ArchiveLimits,
        semantic: SemanticLimits,
    ) -> ReadResult<Self> {
        let (object_index, total_objects) =
            build_object_index(&components, semantic.max_objects())?;
        let package = Self {
            state: Arc::new(State {
                source: PhysicalSource::Semantic(components),
                options: ReadOptions::new(archive, semantic),
                object_index,
                total_objects,
                semantic: OnceCell::new(),
                #[cfg(test)]
                semantic_decode_attempts: AtomicUsize::new(0),
                #[cfg(all(test, feature = "internal-iwork-source"))]
                source_classification_attempts: 0,
            }),
        };
        package.root_show_identifier()?;
        Ok(package)
    }

    /// Capture another handle to the same immutable parsed package.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Start a focused immutable slide-state edit from this package snapshot.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Borrow the original package bytes without normalizing unknown content.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.state.source.source_bytes()
    }

    /// Return the checked physical limits used when this package was parsed.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.options.archive()
    }

    /// Return both checked resource profiles retained by this package.
    #[must_use]
    pub fn read_options(&self) -> ReadOptions {
        self.state.options
    }

    /// Return the checked semantic limits used for lazy projection.
    #[must_use]
    pub fn semantic_limits(&self) -> SemanticLimits {
        self.state.options.semantic()
    }

    /// Return the count of parsed native IWA components.
    #[must_use]
    pub fn component_count(&self) -> usize {
        self.state.source.components().len()
    }

    /// Iterate normalized names for all parsed native IWA components.
    pub fn component_names(&self) -> impl Iterator<Item = &str> {
        self.state
            .source
            .components()
            .iter()
            .map(litchi_iwa_archive::Component::name)
    }

    /// Extract reachable textual content in semantic presentation order.
    ///
    /// Only storages reached through the Keynote show/slide graph participate;
    /// unrelated native messages are never speculatively decoded as text.
    ///
    /// # Errors
    ///
    /// Returns an error only when the parsed package cannot maintain its
    /// validated semantic state.
    pub fn text(&self) -> ReadResult<String> {
        semantic_text(self.show()?)
    }

    /// Borrow semantic slides in presentation order without reparsing the package.
    ///
    /// # Errors
    ///
    /// Returns an error when a required Keynote show, slide, drawable, note,
    /// build, or transition payload cannot be decoded.
    pub fn slides(&self) -> ReadResult<&[Slide]> {
        Ok(self.semantic_document()?.slides())
    }

    /// Borrow the decoded semantic show without reparsing the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the native Keynote show cannot be decoded.
    pub fn show(&self) -> ReadResult<&Show> {
        Ok(self.semantic_document()?.show())
    }

    /// Return a cheap archive-free semantic Keynote document snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when the native Keynote show cannot be decoded.
    pub fn semantic_snapshot(&self) -> ReadResult<Document> {
        Ok(self.semantic_document()?.snapshot())
    }

    /// Extract standard presentation metadata from the semantic show and plist.
    ///
    /// # Errors
    ///
    /// Returns an error when the show or a present `Properties.plist` cannot
    /// be decoded under the package's original physical limits.
    pub fn metadata(&self) -> ReadResult<Option<litchi_core::Metadata>> {
        let mut metadata = litchi_core::Metadata {
            application: Some("Keynote".to_owned()),
            ..litchi_core::Metadata::default()
        };
        let mut has_data = true;

        if let Some(title) = self.show()?.title() {
            metadata.title = Some(title.to_owned());
        }

        if let Some(properties) = self.state.source.package().iter().find(|entry| {
            entry.name().rsplit('/').next() == Some("Properties.plist") && !entry.is_opaque()
        }) {
            let value = plist::Value::from_reader(Cursor::new(properties.data()))?;
            let dictionary = value.as_dictionary().ok_or_else(|| {
                ReadError::InvalidFormat("Keynote Properties.plist is not a dictionary".to_owned())
            })?;
            if metadata.title.is_none() {
                metadata.title = property(dictionary, "Title")
                    .or_else(|| property(dictionary, "kDocumentTitleKey"));
            }
            metadata.author = property(dictionary, "Author")
                .or_else(|| property(dictionary, "kDocumentAuthorKey"))
                .or_else(|| property(dictionary, "kSFWPAuthorPropertyKey"));
            metadata.keywords = property(dictionary, "Keywords");
            metadata.description = property(dictionary, "Comments");
            metadata.revision =
                property(dictionary, "revision").or_else(|| property(dictionary, "buildVersion"));
            metadata.content_status = property(dictionary, "fileFormatVersion")
                .map(|version| format!("Keynote Format Version {version}"));
            has_data = true;
        }

        Ok(has_data.then_some(metadata))
    }

    /// Validate the retained package root and all lazily decoded semantics.
    ///
    /// # Errors
    ///
    /// Returns an error when the package root or required semantic references
    /// are missing or malformed.
    pub fn validate(&self) -> ReadResult<()> {
        self.root_show_identifier()?;
        self.semantic_document()?;
        Ok(())
    }

    /// Return package measurements after resolving the semantic slide tree.
    ///
    /// # Errors
    ///
    /// Returns an error when the Keynote semantic snapshot cannot be decoded.
    pub fn stats(&self) -> ReadResult<Stats> {
        Ok(Stats {
            total_objects: self.state.total_objects,
            slide_count: self.slides()?.len(),
        })
    }

    fn semantic_document(&self) -> ReadResult<&Document> {
        self.state
            .semantic
            .get_or_try_init(|| Ok(Document::from_show(self.decode_show()?)))
    }

    fn root_show_identifier(&self) -> ReadResult<u64> {
        let mut roots = self
            .state
            .source
            .components()
            .iter()
            .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
        let root = roots.next().ok_or_else(|| {
            ReadError::InvalidFormat("missing Index/Document.iwa component".to_owned())
        })?;
        if roots.next().is_some() {
            return Err(ReadError::InvalidFormat(
                "package contains multiple Document.iwa components".to_owned(),
            ));
        }
        let object = root.archive().object(1).ok_or_else(|| {
            ReadError::InvalidFormat("Keynote root object 1 is missing".to_owned())
        })?;
        let payload = unique_payload(
            &object.messages,
            &[DOCUMENT_MESSAGE_TYPE],
            "Keynote root document",
        )?;
        decode_root_show_identifier(payload, self.semantic_wire_limits()?)
    }

    fn decode_show(&self) -> ReadResult<Show> {
        #[cfg(test)]
        {
            self.state
                .semantic_decode_attempts
                .fetch_add(1, Ordering::Relaxed);
            // Make check/decode/set implementations reliably overlap in the
            // concurrency regression. The production build has no delay.
            std::thread::sleep(std::time::Duration::from_millis(10));
        }

        let show_identifier = self.root_show_identifier()?;
        let mut builder = Show::builder();
        if show_identifier == 0 {
            return Ok(builder.build());
        }

        let mut budget = SemanticBudget::new(self.semantic_limits());
        budget.charge_references(1, SemanticPath::Show)?;
        let show_object = self.required_object(show_identifier, "Keynote show")?;
        let payload = unique_payload(&show_object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")?;
        let preflight_slide_count =
            preflight_show(payload, self.semantic_wire_limits()?, &mut budget)?;
        let show = decode_show_snapshot(
            payload,
            self.semantic_limits().max_slides(),
            self.semantic_wire_limits()?,
        )?;
        if show.slide_node_identifiers().len() != preflight_slide_count {
            return Err(ReadError::Decode(
                "Keynote show slide count disagrees with wire preflight".to_owned(),
            ));
        }
        let records = self.slide_records(show.slide_node_identifiers(), &mut budget)?;
        builder
            .try_reserve_slides(records.len())
            .map_err(|_error| ReadError::Allocation {
                resource: "Keynote semantic slides",
                amount: records.len(),
            })?;
        builder.set_settings(settings_from_show(&show)?);
        if let Some(storage) =
            self.optional_text_storage(show_object, &mut budget, SemanticPath::ShowTitle)?
            && !storage.is_empty()
        {
            builder.set_title(Some(storage.into_text()));
        }
        for (index, slide) in records.into_iter().enumerate() {
            let object = self.required_object(slide.slide_identifier, "Keynote slide")?;
            builder.push_slide(self.parse_slide(index, object, slide.is_skipped, &mut budget)?);
        }
        Ok(builder.build())
    }

    fn slide_records(
        &self,
        node_identifiers: &[u64],
        budget: &mut SemanticBudget,
    ) -> ReadResult<Vec<SlideRecord>> {
        let slide_count = node_identifiers.len();
        let maximum = self.semantic_limits().max_slides();
        if slide_count > maximum {
            return Err(ReadError::SemanticLimit {
                kind: SemanticLimitKind::Slides,
                observed: slide_count,
                maximum,
                path: SemanticPath::Show,
            });
        }
        let mut records = Vec::new();
        records
            .try_reserve_exact(slide_count)
            .map_err(|_error| ReadError::Allocation {
                resource: "Keynote slide records",
                amount: slide_count,
            })?;
        let wire_limits = self.semantic_wire_limits()?;
        for (index, &node_identifier) in node_identifiers.iter().enumerate() {
            let node_object = self.required_object(node_identifier, "Keynote slide node")?;
            let node_payload = unique_payload(
                &node_object.messages,
                &[SLIDE_NODE_MESSAGE_TYPE],
                "Keynote slide node",
            )?;
            let is_skipped =
                strict_slide_node_skipped(node_payload, wire_limits).map_err(|error| {
                    map_wire_preflight_error(
                        error,
                        "Keynote slide node",
                        SemanticPath::Slide { index },
                    )
                })?;
            let node_archive: kn::SlideNodeArchive =
                decode_message(node_payload, "Keynote slide node")?;
            if node_archive.is_skipped != is_skipped {
                return Err(ReadError::Decode(
                    "Keynote slide-node skip state disagrees with its wire value".to_owned(),
                ));
            }
            let slide = node_archive.slide.ok_or_else(|| {
                ReadError::InvalidFormat(
                    "Keynote slide node has no required slide reference".to_owned(),
                )
            })?;
            budget.charge_references(1, SemanticPath::Slide { index })?;
            records.push(SlideRecord {
                node_identifier,
                slide_identifier: slide.identifier,
                is_skipped,
            });
        }
        Ok(records)
    }

    fn slide_record_at(&self, index: usize) -> ReadResult<Option<SlideRecord>> {
        let show_identifier = self.root_show_identifier()?;
        if show_identifier == 0 {
            return Ok(None);
        }
        let show_object = self
            .object(show_identifier)
            .ok_or_else(|| ReadError::Decode("Keynote show object is missing".to_owned()))?;
        let payload = unique_payload(&show_object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")?;
        let mut budget = SemanticBudget::new(self.semantic_limits());
        budget.charge_references(1, SemanticPath::Show)?;
        let preflight_slide_count =
            preflight_show(payload, self.semantic_wire_limits()?, &mut budget)?;
        let show = decode_show_snapshot(
            payload,
            self.semantic_limits().max_slides(),
            self.semantic_wire_limits()?,
        )?;
        if show.slide_node_identifiers().len() != preflight_slide_count {
            return Err(ReadError::Decode(
                "Keynote show slide count disagrees with wire preflight".to_owned(),
            ));
        }
        Ok(self
            .slide_records(show.slide_node_identifiers(), &mut budget)?
            .get(index)
            .copied())
    }

    fn parse_slide(
        &self,
        index: usize,
        object: &ArchiveObject,
        is_skipped: bool,
        budget: &mut SemanticBudget,
    ) -> ReadResult<Slide> {
        let payload = unique_payload(&object.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")?;
        let preflight = preflight_slide(payload, self.semantic_wire_limits()?, budget, index)?;
        let slide: kn::SlideArchive = decode_message(payload, "Keynote slide")?;
        if slide.builds.len() != preflight.builds
            || slide.owned_drawables.len() != preflight.owned_drawables
        {
            return Err(ReadError::Decode(
                "Keynote slide reference counts disagree with wire preflight".to_owned(),
            ));
        }
        let mut builder = Slide::builder(index);
        builder.set_skipped(is_skipped);
        builder
            .try_reserve_builds(preflight.builds)
            .map_err(|_error| ReadError::Allocation {
                resource: "Keynote semantic builds",
                amount: preflight.builds,
            })?;
        let storage_capacity = slide
            .owned_drawables
            .len()
            .saturating_add(1)
            .min(budget.remaining_text_storages());
        builder
            .try_reserve_text_storages(storage_capacity)
            .map_err(|_error| ReadError::Allocation {
                resource: "Keynote slide text storages",
                amount: storage_capacity,
            })?;

        if let Some(name) = slide.name.filter(|name| !name.is_empty()) {
            builder.set_name(Some(name));
        }
        for (build_index, build) in slide.builds.iter().enumerate() {
            builder.push_build(self.extract_build(
                build.identifier,
                budget,
                SemanticPath::SlideBuild {
                    slide: index,
                    index: build_index,
                },
            )?);
        }
        builder.set_transition(Some(Self::transition(&slide.transition)?));

        let title = slide
            .title_placeholder
            .as_ref()
            .map(|reference| reference.identifier);
        let body = slide
            .body_placeholder
            .as_ref()
            .map(|reference| reference.identifier);
        if let Some(identifier) = title
            && let Some(storage) =
                self.drawable_storage(identifier, true, budget, SemanticPath::SlideTitle { index })?
            && !storage.is_empty()
        {
            builder.set_title(Some(storage.into_text()));
        }
        if let Some(identifier) = body
            && let Some(storage) =
                self.drawable_storage(identifier, true, budget, SemanticPath::SlideBody { index })?
            && !storage.is_empty()
        {
            builder.push_text_storage(storage);
        }
        for (drawable_index, drawable) in slide.owned_drawables.iter().enumerate() {
            if Some(drawable.identifier) == title || Some(drawable.identifier) == body {
                continue;
            }
            if let Some(storage) = self.drawable_storage(
                drawable.identifier,
                false,
                budget,
                SemanticPath::SlideDrawable {
                    slide: index,
                    index: drawable_index,
                },
            )? && !storage.is_empty()
            {
                builder.push_text_storage(storage);
            }
        }
        if let Some(note) = slide.note {
            let notes_text =
                self.notes_text(note.identifier, budget, SemanticPath::SlideNotes { index })?;
            if !notes_text.is_empty() {
                builder.set_notes(Some(notes_text));
            }
        }
        Ok(builder.build())
    }

    fn extract_build(
        &self,
        identifier: u64,
        budget: &mut SemanticBudget,
        path: SemanticPath,
    ) -> ReadResult<Build> {
        let object = self.required_object(identifier, "Keynote build")?;
        let payload = unique_payload(&object.messages, &[BUILD_MESSAGE_TYPE], "Keynote build")?;
        preflight_build(payload, self.semantic_wire_limits()?, budget, path)?;
        let build: kn::BuildArchive = decode_message(payload, "Keynote build")?;
        let animation = AnimationType::from_identifier(&build.delivery).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote build identifier: {error}"))
        })?;
        #[allow(
            deprecated,
            reason = "native Keynote schemas retain compatibility duration fields"
        )]
        let raw_duration = build
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or(build.attributes.database_duration)
            .or(build.duration)
            .unwrap_or(0.0);
        let duration = Seconds::new(raw_duration).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote build duration: {error}"))
        })?;
        Ok(Build::new(animation, duration))
    }

    fn transition(transition: &kn::TransitionArchive) -> ReadResult<Transition> {
        #[allow(
            deprecated,
            reason = "native Keynote schemas retain compatibility duration fields"
        )]
        let raw_duration = transition
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or(transition.attributes.database_duration)
            .unwrap_or(0.0);
        let duration = Seconds::new(raw_duration).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote transition duration: {error}"))
        })?;
        #[allow(
            deprecated,
            reason = "native Keynote schemas retain compatibility effect fields"
        )]
        let identifier = transition
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.effect.as_deref())
            .or(transition.attributes.database_effect.as_deref());
        let effect = identifier.map_or(Ok(Effect::None), |value| {
            Effect::from_identifier(value).map_err(|error| {
                ReadError::Decode(format!("invalid Keynote transition effect: {error}"))
            })
        })?;
        Ok(Transition::new(effect, duration))
    }

    fn drawable_storage(
        &self,
        identifier: u64,
        required: bool,
        budget: &mut SemanticBudget,
        path: SemanticPath,
    ) -> ReadResult<Option<Storage>> {
        let drawable = self.required_object(identifier, "Keynote drawable")?;
        let placeholder = optional_unique_payload(
            &drawable.messages,
            &[PLACEHOLDER_MESSAGE_TYPE],
            "Keynote drawable placeholder",
        )?;
        let shape = optional_unique_payload(
            &drawable.messages,
            &[SHAPE_INFO_MESSAGE_TYPE],
            "Keynote drawable shape",
        )?;
        if placeholder.is_some() && shape.is_some() {
            return Err(ReadError::InvalidFormat(
                "Keynote drawable contains ambiguous text owners".to_owned(),
            ));
        }

        let storage_reference = if let Some(payload) = placeholder {
            preflight_placeholder(payload, self.semantic_wire_limits()?, path)?;
            let value: kn::PlaceholderArchive =
                decode_message(payload, "Keynote drawable placeholder")?;
            value.super_.owned_storage
        } else if let Some(payload) = shape {
            preflight_required_length_delimited_field(
                payload,
                self.semantic_wire_limits()?,
                1,
                "Keynote shape base archive",
                path,
            )?;
            let value: tswp::ShapeInfoArchive = decode_message(payload, "Keynote drawable shape")?;
            value.owned_storage
        } else {
            None
        };

        let Some(reference) = storage_reference else {
            if required {
                return Err(ReadError::InvalidFormat(
                    "Keynote drawable has no required text-storage reference".to_owned(),
                ));
            }
            return Ok(None);
        };
        budget.charge_references(1, path)?;
        let storage = self.required_object(reference.identifier, "Keynote drawable storage")?;
        self.required_text_storage(storage, budget, path).map(Some)
    }

    fn notes_text(
        &self,
        identifier: u64,
        budget: &mut SemanticBudget,
        path: SemanticPath,
    ) -> ReadResult<String> {
        let note_object = self.required_object(identifier, "Keynote speaker note")?;
        let payload = unique_payload(
            &note_object.messages,
            &[NOTE_MESSAGE_TYPE],
            "Keynote speaker note",
        )?;
        preflight_required_length_delimited_field(
            payload,
            self.semantic_wire_limits()?,
            1,
            "Keynote speaker-note storage reference",
            path,
        )?;
        let note: kn::NoteArchive = decode_message(payload, "Keynote speaker note")?;
        let storage = self.required_object(
            note.contained_storage.identifier,
            "Keynote speaker-note storage",
        )?;
        budget.charge_references(1, path)?;
        Ok(self
            .required_text_storage(storage, budget, path)?
            .into_text())
    }

    fn required_text_storage(
        &self,
        object: &ArchiveObject,
        budget: &mut SemanticBudget,
        path: SemanticPath,
    ) -> ReadResult<Storage> {
        let payload = unique_payload(
            &object.messages,
            &[STORAGE_MESSAGE_TYPE],
            "Keynote text storage",
        )?;
        budget.decode_storage(payload, self.limits(), path)
    }

    fn optional_text_storage(
        &self,
        object: &ArchiveObject,
        budget: &mut SemanticBudget,
        path: SemanticPath,
    ) -> ReadResult<Option<Storage>> {
        optional_unique_payload(
            &object.messages,
            &[STORAGE_MESSAGE_TYPE],
            "Keynote text storage",
        )?
        .map(|payload| budget.decode_storage(payload, self.limits(), path))
        .transpose()
    }

    fn required_object(
        &self,
        identifier: u64,
        context: &'static str,
    ) -> ReadResult<&ArchiveObject> {
        self.object(identifier)
            .ok_or_else(|| ReadError::InvalidFormat(format!("{context} object is missing")))
    }

    fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        let locator = self
            .state
            .object_index
            .binary_search_by_key(&identifier, |locator| locator.identifier)
            .ok()
            .map(|index| self.state.object_index[index])?;
        self.state
            .source
            .components()
            .get_index(locator.component)?
            .archive()
            .objects
            .get(locator.object)
    }

    fn wire_limits(&self) -> litchi_iwa_common::Result<WireLimits> {
        let maximum = self
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?
            .max_message_bytes();
        WireLimits::default()
            .with_input_bytes(maximum)?
            .with_output_bytes(maximum)
    }

    fn semantic_wire_limits(&self) -> ReadResult<WireLimits> {
        self.wire_limits().map_err(|error| {
            map_wire_preflight_error(error, "Keynote semantic payload", SemanticPath::Package)
        })
    }
}

/// Consume a prepared Keynote source into an archive-free semantic document.
///
/// This unstable coordinator handoff releases exact package bytes and logical
/// sidecars before the lazy Keynote graph is projected. It supports both ZIP
/// inputs and app-authored package directories because only the retained IWA
/// component snapshot crosses this boundary.
///
/// # Errors
///
/// Returns [`ReadError`] when the source belongs to another application or
/// its Keynote graph is malformed or exceeds the supplied semantic limits.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __semantic_document_from_prepared_source(
    source: PreparedSource,
    semantic: SemanticLimits,
) -> ReadResult<Document> {
    if source.format() != Format::Keynote {
        return Err(ReadError::NotKeynote);
    }
    let (components, archive) = source.__into_components();
    Package::from_classified_components(components, archive, semantic)?.semantic_snapshot()
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct SlidePreflight {
    builds: usize,
    owned_drawables: usize,
}

#[derive(Debug, Clone, Copy)]
struct SemanticBudget {
    limits: SemanticLimits,
    references: usize,
    text_storages: usize,
    text_fragments: usize,
    text_bytes: usize,
}

impl SemanticBudget {
    const fn new(limits: SemanticLimits) -> Self {
        Self {
            limits,
            references: 0,
            text_storages: 0,
            text_fragments: 0,
            text_bytes: 0,
        }
    }

    fn charge_references(&mut self, amount: usize, path: SemanticPath) -> ReadResult<()> {
        self.references = checked_semantic_charge(
            self.references,
            amount,
            SemanticLimitKind::References,
            self.limits.max_references(),
            path,
        )?;
        Ok(())
    }

    fn charge_text(&mut self, amount: usize, path: SemanticPath) -> ReadResult<()> {
        self.text_bytes = checked_semantic_charge(
            self.text_bytes,
            amount,
            SemanticLimitKind::TextBytes,
            self.limits.max_text_bytes(),
            path,
        )?;
        Ok(())
    }

    fn charge_fragments(&mut self, amount: usize, path: SemanticPath) -> ReadResult<()> {
        self.text_fragments = checked_semantic_charge(
            self.text_fragments,
            amount,
            SemanticLimitKind::TextFragments,
            self.limits.max_text_fragments(),
            path,
        )?;
        Ok(())
    }

    const fn remaining_text_storages(&self) -> usize {
        self.limits
            .max_text_storages()
            .saturating_sub(self.text_storages)
    }

    fn decode_storage(
        &mut self,
        payload: &[u8],
        archive: ArchiveLimits,
        path: SemanticPath,
    ) -> ReadResult<Storage> {
        let storage_count = checked_semantic_charge(
            self.text_storages,
            1,
            SemanticLimitKind::TextStorages,
            self.limits.max_text_storages(),
            path,
        )?;

        let remaining_text = self.limits.max_text_bytes() - self.text_bytes;
        let remaining_fragments = self.limits.max_text_fragments() - self.text_fragments;
        let core = archive
            .effective_archive_limits()
            .map_err(ReadError::Archive)?;
        let text_limits = TextWireLimits::new(
            core.max_message_bytes()
                .min(TextWireLimits::MAX_MESSAGE_BYTES),
            DEFAULT_MAX_TEXT_FIELDS.min(TextWireLimits::MAX_FIELDS),
            remaining_fragments
                .min(DEFAULT_MAX_TEXT_FRAGMENTS)
                .clamp(1, TextWireLimits::MAX_FRAGMENTS),
            remaining_text.clamp(1, TextWireLimits::MAX_TEXT_BYTES),
        )
        .map_err(|_error| ReadError::TextStorage {
            reason: TextStorageFailure::Projection,
            path,
        })?;
        let storage = litchi_iwa_text_wire::from_bytes_with_limits(payload, text_limits)
            .map_err(|error| self.map_text_error(&error, text_limits, path))?;
        self.charge_fragments(storage.runs().len(), path)?;
        self.charge_text(storage.len(), path)?;
        self.text_storages = storage_count;
        Ok(storage)
    }

    fn map_text_error(
        &self,
        error: &TextWireError,
        effective: TextWireLimits,
        path: SemanticPath,
    ) -> ReadError {
        match error {
            TextWireError::TooManyFragments { actual, limit } => ReadError::SemanticLimit {
                kind: SemanticLimitKind::TextFragments,
                observed: self.text_fragments.saturating_add(*actual),
                maximum: self.text_fragments.saturating_add(*limit),
                path,
            },
            TextWireError::TooManyTextBytes { actual, limit } => ReadError::SemanticLimit {
                kind: SemanticLimitKind::TextBytes,
                observed: self.text_bytes.saturating_add(*actual),
                maximum: self.text_bytes.saturating_add(*limit),
                path,
            },
            TextWireError::TextLengthOverflow => ReadError::SemanticLimit {
                kind: SemanticLimitKind::TextBytes,
                observed: usize::MAX,
                maximum: self.text_bytes.saturating_add(effective.max_text_bytes()),
                path,
            },
            TextWireError::InvalidUtf8 { .. } => ReadError::TextStorage {
                reason: TextStorageFailure::InvalidUtf8,
                path,
            },
            TextWireError::WrongTextWireType { .. } => ReadError::TextStorage {
                reason: TextStorageFailure::WrongWireType,
                path,
            },
            TextWireError::ProjectionDecode { .. }
            | TextWireError::ProjectionMismatch { .. }
            | TextWireError::ProjectionTextLengthMismatch { .. }
            | TextWireError::InvalidLimit { .. } => ReadError::TextStorage {
                reason: TextStorageFailure::Projection,
                path,
            },
            TextWireError::Storage(_) => ReadError::TextStorage {
                reason: TextStorageFailure::InvalidRanges,
                path,
            },
            TextWireError::Common(litchi_iwa_common::Error::LimitExceeded {
                kind,
                observed,
                limit,
            }) => ReadError::PayloadLimit {
                kind: payload_limit_kind(*kind),
                observed: *observed,
                maximum: *limit,
                path,
            },
            TextWireError::Common(litchi_iwa_common::Error::Allocation { amount, .. }) => {
                ReadError::Allocation {
                    resource: "Keynote semantic text storage",
                    amount: *amount,
                }
            },
            TextWireError::Common(_) | _ => ReadError::TextStorage {
                reason: TextStorageFailure::MalformedWire,
                path,
            },
        }
    }
}

fn preflight_document(payload: &[u8], wire_limits: WireLimits) -> ReadResult<()> {
    let mut show_fields = 0usize;
    let mut show_identifier_fields = 0usize;
    let mut super_fields = 0usize;
    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field();
        match (visit.path(), field.number()) {
            ([], 2) => {
                require_unique_length_delimited(
                    field,
                    &mut show_fields,
                    "Keynote document show reference",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], 3) => {
                require_unique_length_delimited(
                    field,
                    &mut super_fields,
                    "Keynote document base archive",
                )?;
                Ok(WireDescent::Skip)
            },
            ([2], 1) => {
                require_unique_uint64(
                    field,
                    &mut show_identifier_fields,
                    "Keynote document show identifier",
                )?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    })
    .map_err(|error| map_wire_preflight_error(error, "Keynote document", SemanticPath::Package))?;
    if show_fields != 1 || show_identifier_fields != 1 || super_fields != 1 {
        return Err(ReadError::InvalidFormat(
            "Keynote document is missing a unique required envelope field".to_owned(),
        ));
    }
    Ok(())
}

fn decode_root_show_identifier(payload: &[u8], wire_limits: WireLimits) -> ReadResult<u64> {
    preflight_document(payload, wire_limits)?;
    let recursion_limit = u32::try_from(wire_limits.max_nesting()).map_err(|_error| {
        ReadError::InvalidFormat("Keynote root nesting limit does not fit u32".to_owned())
    })?;
    keynote_document_codec::decode_show_identifier(
        payload,
        keynote_document_codec::DecodeOptions::new(payload.len(), recursion_limit),
    )
    .map_err(|error| {
        ReadError::InvalidFormat(format!(
            "Keynote root document projection is malformed: {error}"
        ))
    })
}

fn decode_show_snapshot(
    payload: &[u8],
    max_slide_references: usize,
    wire_limits: WireLimits,
) -> ReadResult<keynote_show_codec::ShowSnapshot> {
    let recursion_limit = u32::try_from(wire_limits.max_nesting()).map_err(|_error| {
        ReadError::InvalidFormat("Keynote show nesting limit does not fit u32".to_owned())
    })?;
    keynote_show_codec::decode_show(
        payload,
        keynote_show_codec::DecodeOptions::new(
            payload.len(),
            max_slide_references,
            recursion_limit,
        ),
    )
    .map_err(|error| {
        if let Some((observed, maximum)) = error.slide_reference_limit_values() {
            ReadError::SemanticLimit {
                kind: SemanticLimitKind::Slides,
                observed,
                maximum,
                path: SemanticPath::Show,
            }
        } else if let Some(amount) = error.allocation_amount() {
            ReadError::Allocation {
                resource: "Keynote show slide references",
                amount,
            }
        } else {
            ReadError::InvalidFormat(format!("Keynote show projection is malformed: {error}"))
        }
    })
}

fn decode_show_settings_snapshot(
    payload: &[u8],
    max_slide_references: usize,
    wire_limits: WireLimits,
) -> ReadResult<keynote_show_codec::SettingsSnapshot> {
    let recursion_limit = u32::try_from(wire_limits.max_nesting()).map_err(|_error| {
        ReadError::InvalidFormat("Keynote show nesting limit does not fit u32".to_owned())
    })?;
    keynote_show_codec::decode_settings(
        payload,
        keynote_show_codec::DecodeOptions::new(
            payload.len(),
            max_slide_references,
            recursion_limit,
        ),
    )
    .map_err(|error| {
        if let Some((observed, maximum)) = error.slide_reference_limit_values() {
            ReadError::SemanticLimit {
                kind: SemanticLimitKind::Slides,
                observed,
                maximum,
                path: SemanticPath::Show,
            }
        } else if let Some(amount) = error.allocation_amount() {
            ReadError::Allocation {
                resource: "Keynote show settings projection",
                amount,
            }
        } else {
            ReadError::InvalidFormat(format!(
                "Keynote show settings projection is malformed: {error}"
            ))
        }
    })
}

fn preflight_show(
    payload: &[u8],
    wire_limits: WireLimits,
    budget: &mut SemanticBudget,
) -> ReadResult<usize> {
    const UI_STATE_FIELD: u32 = 1;
    const THEME_FIELD: u32 = 2;
    const SLIDE_TREE_FIELD: u32 = 3;
    const SIZE_FIELD: u32 = 4;
    const STYLESHEET_FIELD: u32 = 5;
    const SLIDE_NUMBERS_FIELD: u32 = 6;
    const RECORDING_FIELD: u32 = 7;
    const LOOP_FIELD: u32 = 8;
    const MODE_FIELD: u32 = 9;
    const TRANSITION_DELAY_FIELD: u32 = 10;
    const BUILD_DELAY_FIELD: u32 = 11;
    const IDLE_ACTIVE_FIELD: u32 = 15;
    const IDLE_DELAY_FIELD: u32 = 16;
    const SOUNDTRACK_FIELD: u32 = 17;
    const PLAY_ON_OPEN_FIELD: u32 = 18;
    const SLIDE_LIST_FIELD: u32 = 19;

    let mut ui_state_fields = 0usize;
    let mut theme_fields = 0usize;
    let mut slide_tree_fields = 0usize;
    let mut size_fields = 0usize;
    let mut stylesheet_fields = 0usize;
    let mut recording_fields = 0usize;
    let mut soundtrack_fields = 0usize;
    let mut slide_list_fields = 0usize;
    let mut slide_number_fields = 0usize;
    let mut loop_fields = 0usize;
    let mut mode_fields = 0usize;
    let mut transition_delay_fields = 0usize;
    let mut build_delay_fields = 0usize;
    let mut idle_active_fields = 0usize;
    let mut idle_delay_fields = 0usize;
    let mut play_on_open_fields = 0usize;
    let mut root_slide_node_fields = 0usize;
    let mut slides = 0usize;
    let maximum = budget.limits.max_slides();
    let result = preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field();
        match (visit.path(), field.number()) {
            ([], UI_STATE_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut ui_state_fields,
                    wire_limits,
                    "Keynote show UI state",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], THEME_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut theme_fields,
                    wire_limits,
                    "Keynote show theme",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], SLIDE_TREE_FIELD) => {
                require_unique_canonical_length_delimited(
                    field,
                    &mut slide_tree_fields,
                    "Keynote show slide tree",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], SIZE_FIELD) => {
                require_unique_canonical_length_delimited(
                    field,
                    &mut size_fields,
                    "Keynote show size",
                )?;
                validate_size(field.payload(), wire_limits)?;
                Ok(WireDescent::Descend)
            },
            ([], STYLESHEET_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut stylesheet_fields,
                    wire_limits,
                    "Keynote show stylesheet",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], RECORDING_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut recording_fields,
                    wire_limits,
                    "Keynote show recording",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], SOUNDTRACK_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut soundtrack_fields,
                    wire_limits,
                    "Keynote show soundtrack",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], SLIDE_LIST_FIELD) => {
                validate_unique_reference(
                    field,
                    &mut slide_list_fields,
                    wire_limits,
                    "Keynote show slide list",
                )?;
                Ok(WireDescent::Descend)
            },
            ([], SLIDE_NUMBERS_FIELD) => {
                require_unique_bool(field, &mut slide_number_fields, "Keynote slide numbers")?;
                Ok(WireDescent::Skip)
            },
            ([], LOOP_FIELD) => {
                require_unique_bool(field, &mut loop_fields, "Keynote show loop state")?;
                Ok(WireDescent::Skip)
            },
            ([], MODE_FIELD) => {
                require_unique_canonical_int32(field, &mut mode_fields, "Keynote show mode")?;
                Ok(WireDescent::Skip)
            },
            ([], TRANSITION_DELAY_FIELD) => {
                require_unique_fixed64(
                    field,
                    &mut transition_delay_fields,
                    "Keynote autoplay transition delay",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], BUILD_DELAY_FIELD) => {
                require_unique_fixed64(
                    field,
                    &mut build_delay_fields,
                    "Keynote autoplay build delay",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], IDLE_ACTIVE_FIELD) => {
                require_unique_bool(field, &mut idle_active_fields, "Keynote idle timer state")?;
                Ok(WireDescent::Skip)
            },
            ([], IDLE_DELAY_FIELD) => {
                require_unique_fixed64(field, &mut idle_delay_fields, "Keynote idle timer delay")?;
                Ok(WireDescent::Skip)
            },
            ([], PLAY_ON_OPEN_FIELD) => {
                require_unique_bool(
                    field,
                    &mut play_on_open_fields,
                    "Keynote play-on-open state",
                )?;
                Ok(WireDescent::Skip)
            },
            ([SLIDE_TREE_FIELD], 1) => {
                validate_unique_reference(
                    field,
                    &mut root_slide_node_fields,
                    wire_limits,
                    "Keynote root slide node",
                )?;
                Ok(WireDescent::Descend)
            },
            ([SLIDE_TREE_FIELD], 2) => {
                require_canonical_length_delimited(field, "Keynote show slide reference")?;
                increment_wire_count(&mut slides, "Keynote show slides")?;
                if slides > maximum {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "Keynote semantic slide limit reached during preflight".to_owned(),
                    ));
                }
                // The Buffa projection validates the reference and owns the
                // sole slide-order buffer. Still descend so aggregate wire
                // bytes, fields, nesting, and work include every reference.
                Ok(WireDescent::Descend)
            },
            _ => Ok(WireDescent::Skip),
        }
    });
    if slides > maximum {
        return Err(ReadError::SemanticLimit {
            kind: SemanticLimitKind::Slides,
            observed: slides,
            maximum,
            path: SemanticPath::Show,
        });
    }
    result.map_err(|error| map_wire_preflight_error(error, "Keynote show", SemanticPath::Show))?;
    if theme_fields != 1 || slide_tree_fields != 1 || size_fields != 1 || stylesheet_fields != 1 {
        return Err(ReadError::InvalidFormat(
            "Keynote show is missing a unique required envelope field".to_owned(),
        ));
    }
    budget.charge_references(slides, SemanticPath::Show)?;
    Ok(slides)
}

fn validate_unique_reference(
    field: WireFieldView<'_>,
    count: &mut usize,
    limits: WireLimits,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_unique_canonical_length_delimited(field, count, context)?;
    validate_reference_payload(field.payload(), limits, context).map(|_identifier| ())
}

fn validate_reference_payload(
    payload: &[u8],
    limits: WireLimits,
    context: &'static str,
) -> litchi_iwa_common::Result<u64> {
    let view = WireView::parse_with_limits(payload, limits)?;
    let mut identifier_fields = 0usize;
    let mut deprecated_type_fields = 0usize;
    let mut deprecated_external_fields = 0usize;
    let mut identifier = None;
    for field in view.fields() {
        match field.number() {
            1 => {
                identifier = Some(require_unique_uint64(
                    field,
                    &mut identifier_fields,
                    context,
                )?);
            },
            2 => {
                require_unique_canonical_int32(field, &mut deprecated_type_fields, context)?;
            },
            3 => {
                require_unique_bool(field, &mut deprecated_external_fields, context)?;
            },
            _ => {},
        }
    }
    identifier.ok_or_else(|| {
        litchi_iwa_common::Error::InvalidFormat(format!("{context} has no required identifier"))
    })
}

fn validate_size(payload: &[u8], limits: WireLimits) -> litchi_iwa_common::Result<()> {
    let view = WireView::parse_with_limits(payload, limits)?;
    let mut width_fields = 0usize;
    let mut height_fields = 0usize;
    for field in view.fields() {
        match field.number() {
            1 => require_unique_fixed32(field, &mut width_fields, "Keynote show width")?,
            2 => require_unique_fixed32(field, &mut height_fields, "Keynote show height")?,
            _ => {},
        }
    }
    if width_fields != 1 || height_fields != 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(
            "Keynote show size is missing a unique dimension".to_owned(),
        ));
    }
    Ok(())
}

fn preflight_slide(
    payload: &[u8],
    wire_limits: WireLimits,
    budget: &mut SemanticBudget,
    index: usize,
) -> ReadResult<SlidePreflight> {
    let mut style_fields = 0usize;
    let mut builds = 0usize;
    let mut owned_drawables = 0usize;
    let mut title_fields = 0usize;
    let mut body_fields = 0usize;
    let mut note_fields = 0usize;
    let mut transition_fields = 0usize;
    let mut transition_attribute_fields = 0usize;
    let mut animation_attribute_fields = 0usize;
    let mut in_document_fields = 0usize;
    let mut name_bytes = None;
    let mut database_effect_bytes = None;
    let mut animation_effect_bytes = None;

    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field();
        match (visit.path(), field.number()) {
            ([], 1) => {
                require_unique_length_delimited(field, &mut style_fields, "Keynote slide style")?;
                Ok(WireDescent::Skip)
            },
            ([], 2) => {
                require_length_delimited(field, "Keynote slide build reference")?;
                increment_wire_count(&mut builds, "Keynote slide build references")?;
                Ok(WireDescent::Skip)
            },
            ([], 7) => {
                require_length_delimited(field, "Keynote slide drawable reference")?;
                increment_wire_count(&mut owned_drawables, "Keynote slide drawable references")?;
                Ok(WireDescent::Skip)
            },
            ([], 5) => {
                require_unique_length_delimited(
                    field,
                    &mut title_fields,
                    "Keynote slide title reference",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], 6) => {
                require_unique_length_delimited(
                    field,
                    &mut body_fields,
                    "Keynote slide body reference",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], 27) => {
                require_unique_length_delimited(
                    field,
                    &mut note_fields,
                    "Keynote slide note reference",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], 10) => {
                set_unique_payload_len(field, &mut name_bytes, "Keynote slide name")?;
                Ok(WireDescent::Skip)
            },
            ([], 19) => {
                require_unique_bool(
                    field,
                    &mut in_document_fields,
                    "Keynote slide in-document state",
                )?;
                Ok(WireDescent::Skip)
            },
            ([], 4) => {
                require_unique_length_delimited(
                    field,
                    &mut transition_fields,
                    "Keynote slide transition",
                )?;
                Ok(WireDescent::Descend)
            },
            ([4], 2) => {
                require_unique_length_delimited(
                    field,
                    &mut transition_attribute_fields,
                    "Keynote transition attributes",
                )?;
                Ok(WireDescent::Descend)
            },
            ([4, 2], 8) => {
                require_unique_length_delimited(
                    field,
                    &mut animation_attribute_fields,
                    "Keynote transition animation attributes",
                )?;
                Ok(WireDescent::Descend)
            },
            ([4, 2], 2) => {
                set_unique_payload_len(
                    field,
                    &mut database_effect_bytes,
                    "Keynote transition database effect",
                )?;
                Ok(WireDescent::Skip)
            },
            ([4, 2, 8], 2) => {
                set_unique_payload_len(
                    field,
                    &mut animation_effect_bytes,
                    "Keynote transition effect",
                )?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    })
    .map_err(|error| {
        map_wire_preflight_error(error, "Keynote slide", SemanticPath::Slide { index })
    })?;

    if style_fields != 1
        || transition_fields != 1
        || transition_attribute_fields != 1
        || in_document_fields != 1
    {
        return Err(ReadError::InvalidFormat(
            "Keynote slide is missing a unique required envelope field".to_owned(),
        ));
    }
    let references = builds
        .checked_add(owned_drawables)
        .and_then(|value| value.checked_add(title_fields))
        .and_then(|value| value.checked_add(body_fields))
        .and_then(|value| value.checked_add(note_fields))
        .ok_or_else(|| {
            ReadError::InvalidFormat("Keynote slide reference count overflowed".to_owned())
        })?;
    budget.charge_references(references, SemanticPath::Slide { index })?;
    budget.charge_text(name_bytes.unwrap_or(0), SemanticPath::SlideName { index })?;
    budget.charge_text(
        animation_effect_bytes
            .or(database_effect_bytes)
            .unwrap_or(0),
        SemanticPath::SlideTransition { index },
    )?;
    Ok(SlidePreflight {
        builds,
        owned_drawables,
    })
}

fn preflight_build(
    payload: &[u8],
    wire_limits: WireLimits,
    budget: &mut SemanticBudget,
    path: SemanticPath,
) -> ReadResult<()> {
    let mut delivery_bytes = None;
    let mut attribute_fields = 0usize;
    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field();
        if visit.path().is_empty() && field.number() == 2 {
            set_unique_payload_len(field, &mut delivery_bytes, "Keynote build delivery")?;
        } else if visit.path().is_empty() && field.number() == 4 {
            require_unique_length_delimited(
                field,
                &mut attribute_fields,
                "Keynote build attributes",
            )?;
        }
        Ok(WireDescent::Skip)
    })
    .map_err(|error| map_wire_preflight_error(error, "Keynote build", path))?;
    let amount = delivery_bytes.ok_or_else(|| {
        ReadError::InvalidFormat("Keynote build has no unique delivery identifier".to_owned())
    })?;
    if attribute_fields != 1 {
        return Err(ReadError::InvalidFormat(
            "Keynote build has no unique required attributes".to_owned(),
        ));
    }
    budget.charge_text(amount, path)
}

fn preflight_required_length_delimited_field(
    payload: &[u8],
    wire_limits: WireLimits,
    field_number: u32,
    context: &'static str,
    path: SemanticPath,
) -> ReadResult<()> {
    let mut fields = 0usize;
    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        if visit.path().is_empty() && visit.field().number() == field_number {
            require_unique_length_delimited(visit.field(), &mut fields, context)?;
        }
        Ok(WireDescent::Skip)
    })
    .map_err(|error| map_wire_preflight_error(error, context, path))?;
    if fields != 1 {
        return Err(ReadError::InvalidFormat(format!(
            "{context} is missing or duplicated"
        )));
    }
    Ok(())
}

fn preflight_placeholder(
    payload: &[u8],
    wire_limits: WireLimits,
    path: SemanticPath,
) -> ReadResult<()> {
    let mut placeholder_super_fields = 0usize;
    let mut shape_super_fields = 0usize;
    preflight_wire_tree_with_limits(payload, wire_limits, |visit| {
        let field = visit.field();
        match (visit.path(), field.number()) {
            ([], 1) => {
                require_unique_length_delimited(
                    field,
                    &mut placeholder_super_fields,
                    "Keynote placeholder shape owner",
                )?;
                Ok(WireDescent::Descend)
            },
            ([1], 1) => {
                require_unique_length_delimited(
                    field,
                    &mut shape_super_fields,
                    "Keynote placeholder base shape",
                )?;
                Ok(WireDescent::Skip)
            },
            _ => Ok(WireDescent::Skip),
        }
    })
    .map_err(|error| map_wire_preflight_error(error, "Keynote placeholder", path))?;
    if placeholder_super_fields != 1 || shape_super_fields != 1 {
        return Err(ReadError::InvalidFormat(
            "Keynote placeholder is missing a unique required shape envelope".to_owned(),
        ));
    }
    Ok(())
}

fn require_length_delimited(
    field: WireFieldView<'_>,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    if field.wire_type() != 2 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is not length-delimited"
        )));
    }
    Ok(())
}

fn require_canonical_length_delimited(
    field: WireFieldView<'_>,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_length_delimited(field, context)?;
    field.validate_canonical_framing()
}

fn increment_wire_count(count: &mut usize, context: &'static str) -> litchi_iwa_common::Result<()> {
    *count = count.checked_add(1).ok_or_else(|| {
        litchi_iwa_common::Error::InvalidFormat(format!("{context} overflowed usize"))
    })?;
    Ok(())
}

fn require_unique_length_delimited(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_length_delimited(field, context)?;
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    Ok(())
}

fn require_unique_canonical_length_delimited(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_canonical_length_delimited(field, context)?;
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    Ok(())
}

fn require_unique_canonical_int32(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    field.validate_canonical_key()?;
    if field.wire_type() != 0 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is not a varint"
        )));
    }
    let (value, consumed) =
        litchi_iwa_common::decode_varint_from_bytes(field.payload()).map_err(|error| {
            litchi_iwa_common::Error::InvalidFormat(format!(
                "{context} has an invalid varint: {error}"
            ))
        })?;
    if consumed != field.payload().len()
        || litchi_iwa_common::varint::encoded_len(value) != consumed
        || (value > i32::MAX as u64 && value < 0xffff_ffff_8000_0000)
    {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} has noncanonical int32 framing"
        )));
    }
    Ok(())
}

fn require_unique_fixed32(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_unique_fixed(field, count, context, 5, 4)
}

fn require_unique_fixed64(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_unique_fixed(field, count, context, 1, 8)
}

fn require_unique_fixed(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
    wire_type: u8,
    width: usize,
) -> litchi_iwa_common::Result<()> {
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    field.validate_canonical_key()?;
    if field.wire_type() != wire_type || field.payload().len() != width {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} has the wrong fixed-width wire representation"
        )));
    }
    Ok(())
}

fn require_unique_uint64(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<u64> {
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    field.validate_canonical_key()?;
    if field.wire_type() != 0 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is not a varint"
        )));
    }
    let payload = field.payload();
    let (value, consumed) =
        litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
            litchi_iwa_common::Error::InvalidFormat(format!(
                "{context} has an invalid varint: {error}"
            ))
        })?;
    if consumed != payload.len() || litchi_iwa_common::varint::encoded_len(value) != payload.len() {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} has noncanonical varint framing"
        )));
    }
    Ok(value)
}

fn require_unique_bool(
    field: WireFieldView<'_>,
    count: &mut usize,
    context: &'static str,
) -> litchi_iwa_common::Result<bool> {
    increment_wire_count(count, context)?;
    if *count > 1 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    field.validate_canonical_key()?;
    if field.wire_type() != 0 {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is not a varint"
        )));
    }
    let payload = field.payload();
    let (value, consumed) =
        litchi_iwa_common::decode_varint_from_bytes(payload).map_err(|error| {
            litchi_iwa_common::Error::InvalidFormat(format!(
                "{context} has an invalid varint: {error}"
            ))
        })?;
    if consumed != payload.len() || litchi_iwa_common::varint::encoded_len(value) != payload.len() {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} has noncanonical varint framing"
        )));
    }
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is not a canonical Boolean"
        ))),
    }
}

fn set_unique_payload_len(
    field: WireFieldView<'_>,
    slot: &mut Option<usize>,
    context: &'static str,
) -> litchi_iwa_common::Result<()> {
    require_length_delimited(field, context)?;
    if slot.replace(field.payload().len()).is_some() {
        return Err(litchi_iwa_common::Error::InvalidFormat(format!(
            "{context} is duplicated"
        )));
    }
    Ok(())
}

fn map_wire_preflight_error(
    error: litchi_iwa_common::Error,
    context: &'static str,
    path: SemanticPath,
) -> ReadError {
    match error {
        litchi_iwa_common::Error::Allocation { amount, .. } => ReadError::Allocation {
            resource: "Keynote wire preflight",
            amount,
        },
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ReadError::PayloadLimit {
            kind: payload_limit_kind(kind),
            observed,
            maximum: limit,
            path,
        },
        other @ (litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. }) => {
            ReadError::InvalidFormat(format!("{context} wire preflight failed: {other}"))
        },
    }
}

const fn payload_limit_kind(kind: litchi_iwa_common::LimitKind) -> PayloadLimitKind {
    match kind {
        litchi_iwa_common::LimitKind::InputBytes | litchi_iwa_common::LimitKind::OutputBytes => {
            PayloadLimitKind::Bytes
        },
        litchi_iwa_common::LimitKind::Fields
        | litchi_iwa_common::LimitKind::TableRows
        | litchi_iwa_common::LimitKind::TableColumns
        | litchi_iwa_common::LimitKind::TableCells
        | litchi_iwa_common::LimitKind::MaterializedCells => PayloadLimitKind::Fields,
        litchi_iwa_common::LimitKind::Nesting => PayloadLimitKind::Nesting,
        litchi_iwa_common::LimitKind::RewriteWork => PayloadLimitKind::Work,
    }
}

fn checked_semantic_charge(
    current: usize,
    amount: usize,
    kind: SemanticLimitKind,
    maximum: usize,
    path: SemanticPath,
) -> ReadResult<usize> {
    let observed = current
        .checked_add(amount)
        .ok_or(ReadError::SemanticLimit {
            kind,
            observed: usize::MAX,
            maximum,
            path,
        })?;
    if observed > maximum {
        return Err(ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            path,
        });
    }
    Ok(observed)
}

fn build_object_index(
    components: &ComponentCatalog,
    maximum: usize,
) -> ReadResult<(Box<[ObjectLocator]>, usize)> {
    let mut total_objects = 0usize;
    for component in components.iter() {
        total_objects = total_objects
            .checked_add(component.archive().objects.len())
            .ok_or(ReadError::SemanticLimit {
                kind: SemanticLimitKind::Objects,
                observed: usize::MAX,
                maximum,
                path: SemanticPath::Package,
            })?;
        if total_objects > maximum {
            return Err(ReadError::SemanticLimit {
                kind: SemanticLimitKind::Objects,
                observed: total_objects,
                maximum,
                path: SemanticPath::Package,
            });
        }
    }

    let mut index = Vec::new();
    index
        .try_reserve_exact(total_objects)
        .map_err(|_error| ReadError::Allocation {
            resource: "Keynote object index",
            amount: total_objects,
        })?;
    for (component_index, component) in components.iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            if let Some(identifier) = object.archive_info.identifier {
                index.push(ObjectLocator {
                    identifier,
                    component: component_index,
                    object: object_index,
                });
            }
        }
    }
    index.sort_unstable_by_key(|locator| locator.identifier);
    if index
        .windows(2)
        .any(|window| window[0].identifier == window[1].identifier)
    {
        return Err(ReadError::InvalidFormat(
            "Keynote package contains duplicate native object identities".to_owned(),
        ));
    }
    Ok((index.into_boxed_slice(), total_objects))
}

fn unique_payload<'a>(
    messages: &'a [RawMessage],
    message_types: &[u32],
    context: &'static str,
) -> ReadResult<&'a [u8]> {
    optional_unique_payload(messages, message_types, context)?
        .ok_or_else(|| ReadError::InvalidFormat(format!("{context} has no required typed payload")))
}

fn optional_unique_payload<'a>(
    messages: &'a [RawMessage],
    message_types: &[u32],
    context: &'static str,
) -> ReadResult<Option<&'a [u8]>> {
    let mut matches = messages
        .iter()
        .filter(|message| message_types.contains(&message.type_));
    let payload = matches.next().map(|message| message.data.as_slice());
    if matches.next().is_some() {
        return Err(ReadError::Decode(format!(
            "{context} contains duplicate typed payloads"
        )));
    }
    Ok(payload)
}

fn decode_message<M>(payload: &[u8], context: &'static str) -> ReadResult<M>
where
    M: Message + Default,
{
    M::decode(payload)
        .map_err(|_error| ReadError::InvalidFormat(format!("{context} payload is malformed")))
}

fn semantic_text(show: &Show) -> ReadResult<String> {
    let mut parts = 0usize;
    let mut content_bytes = 0usize;
    visit_show_text(show, |text| {
        parts = parts.checked_add(1).ok_or(ReadError::Allocation {
            resource: "Keynote text part count",
            amount: usize::MAX,
        })?;
        content_bytes = content_bytes
            .checked_add(text.len())
            .ok_or(ReadError::Allocation {
                resource: "Keynote extracted text",
                amount: usize::MAX,
            })?;
        Ok(())
    })?;
    let separator_bytes = parts.saturating_sub(1);
    let total_bytes = content_bytes
        .checked_add(separator_bytes)
        .ok_or(ReadError::Allocation {
            resource: "Keynote extracted text",
            amount: usize::MAX,
        })?;
    let mut output = String::new();
    output
        .try_reserve_exact(total_bytes)
        .map_err(|_error| ReadError::Allocation {
            resource: "Keynote extracted text",
            amount: total_bytes,
        })?;
    let mut first = true;
    visit_show_text(show, |text| {
        if !first {
            output.push('\n');
        }
        output.push_str(text);
        first = false;
        Ok(())
    })?;
    Ok(output)
}

fn visit_show_text(show: &Show, mut visit: impl FnMut(&str) -> ReadResult<()>) -> ReadResult<()> {
    if let Some(title) = show.title().filter(|text| !text.is_empty()) {
        visit(title)?;
    }
    for slide in show.slides() {
        if let Some(title) = slide.title().filter(|text| !text.is_empty()) {
            visit(title)?;
        }
        for text in slide.text_content().iter().filter(|text| !text.is_empty()) {
            visit(text)?;
        }
        for storage in slide
            .text_storages()
            .iter()
            .filter(|storage| !storage.is_empty())
        {
            visit(storage.text())?;
        }
        if let Some(notes) = slide.notes().filter(|text| !text.is_empty()) {
            visit(notes)?;
        }
    }
    Ok(())
}

fn strict_slide_node_skipped(data: &[u8], limits: WireLimits) -> litchi_iwa_common::Result<bool> {
    let mut slide_fields = 0usize;
    let mut skipped_fields = 0usize;
    let mut has_builds_fields = 0usize;
    let mut has_transition_fields = 0usize;
    let mut skipped = false;
    preflight_wire_tree_with_limits(data, limits, |visit| {
        let field = visit.field();
        if !visit.path().is_empty() {
            return Ok(WireDescent::Skip);
        }
        match field.number() {
            2 => require_unique_length_delimited(
                field,
                &mut slide_fields,
                "Keynote slide-node slide reference",
            )?,
            4 => {
                skipped = require_unique_bool(
                    field,
                    &mut skipped_fields,
                    "Keynote slide-node skip state",
                )?;
            },
            6 => {
                require_unique_bool(
                    field,
                    &mut has_builds_fields,
                    "Keynote slide-node has-builds state",
                )?;
            },
            7 => {
                require_unique_bool(
                    field,
                    &mut has_transition_fields,
                    "Keynote slide-node has-transition state",
                )?;
            },
            _ => {},
        }
        Ok(WireDescent::Skip)
    })?;
    if slide_fields != 1
        || skipped_fields != 1
        || has_builds_fields != 1
        || has_transition_fields != 1
    {
        return Err(litchi_iwa_common::Error::InvalidFormat(
            "Keynote slide node is missing a unique required envelope field".to_owned(),
        ));
    }
    Ok(skipped)
}

fn read_source(path: &Path, limits: Limits) -> ReadResult<Arc<[u8]>> {
    let mut file = File::open(path)?;
    let length = file.metadata()?.len();
    read_source_with_reported_length(&mut file, length, limits)
}

fn read_source_with_reported_length(
    reader: &mut impl Read,
    reported_length: u64,
    limits: Limits,
) -> ReadResult<Arc<[u8]>> {
    check_input_size(reported_length, limits)?;

    let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
        ReadError::InvalidFormat("Keynote input limit does not fit usize".to_owned())
    })?;
    let capacity = usize::try_from(reported_length).map_err(|_error| {
        ReadError::InvalidFormat("Keynote input length does not fit usize".to_owned())
    })?;
    let mut bytes = Vec::new();
    bytes.try_reserve_exact(capacity).map_err(|_error| {
        ReadError::Archive(litchi_iwa_archive::Error::Allocation {
            resource: "Keynote package input",
            amount: capacity,
        })
    })?;

    let mut buffer = [0u8; 8 * 1024];
    loop {
        let remaining = maximum.checked_sub(bytes.len()).ok_or_else(|| {
            ReadError::InvalidFormat("Keynote input length exceeds usize".to_owned())
        })?;
        if remaining == 0 {
            let mut extra = [0u8; 1];
            if reader.read(&mut extra)? != 0 {
                return Err(input_limit_error(
                    limits.max_input_bytes().saturating_add(1),
                    limits,
                ));
            }
            break;
        }

        let read_limit = remaining.min(buffer.len());
        let read = reader.read(&mut buffer[..read_limit])?;
        if read == 0 {
            break;
        }
        let required = bytes.len().checked_add(read).ok_or_else(|| {
            ReadError::InvalidFormat("Keynote input length exceeds usize".to_owned())
        })?;
        reserve_source_growth(&mut bytes, required, maximum)?;
        bytes.extend_from_slice(&buffer[..read]);
    }

    Ok(bytes.into())
}

fn reserve_source_growth(bytes: &mut Vec<u8>, required: usize, maximum: usize) -> ReadResult<()> {
    if required <= bytes.capacity() {
        return Ok(());
    }

    // A regular file can grow after `metadata`. Retain amortized linear
    // growth without ever requesting capacity beyond the physical ceiling.
    let doubled = bytes.capacity().checked_mul(2).unwrap_or(maximum);
    let target = required.max(doubled).min(maximum);
    let additional = target
        .checked_sub(bytes.len())
        .ok_or_else(|| ReadError::InvalidFormat("Keynote input length exceeds usize".to_owned()))?;
    bytes.try_reserve_exact(additional).map_err(|_error| {
        ReadError::Archive(litchi_iwa_archive::Error::Allocation {
            resource: "Keynote package input",
            amount: target,
        })
    })
}

fn copy_source(bytes: &[u8]) -> ReadResult<Arc<[u8]>> {
    let mut source = Vec::new();
    source.try_reserve_exact(bytes.len()).map_err(|_error| {
        ReadError::Archive(litchi_iwa_archive::Error::Allocation {
            resource: "Keynote package input",
            amount: bytes.len(),
        })
    })?;
    source.extend_from_slice(bytes);
    Ok(source.into())
}

fn check_input_size(size: u64, limits: Limits) -> ReadResult<()> {
    if size > limits.max_input_bytes() {
        return Err(input_limit_error(size, limits));
    }
    Ok(())
}

fn input_limit_error(observed: u64, limits: Limits) -> ReadError {
    ReadError::Archive(litchi_iwa_archive::Error::Limit {
        kind: litchi_iwa_archive::LimitKind::InputBytes,
        observed,
        maximum: limits.max_input_bytes(),
    })
}

fn settings_from_show(show: &keynote_show_codec::ShowSnapshot) -> ReadResult<Settings> {
    settings_from_show_projection(show.size(), show.raw_settings())
}

fn settings_from_show_projection(
    raw_size: keynote_show_codec::RawSize,
    raw: keynote_show_codec::RawSettings,
) -> ReadResult<Settings> {
    let semantic_size = Size::new(raw_size.width(), raw_size.height())
        .map_err(|error| ReadError::Decode(format!("invalid Keynote show size: {error}")))?;
    let mut settings = Settings::new(semantic_size);
    settings.set_slide_numbers_visible(raw.slide_numbers_visible());
    settings.set_loop_presentation(raw.loop_presentation());
    settings
        .set_mode(raw.mode_raw().map(Mode::from_raw))
        .map_err(|error| ReadError::Decode(format!("invalid Keynote show mode: {error}")))?;
    settings.set_autoplay_transition_delay(
        raw.autoplay_transition_delay()
            .map(Seconds::new)
            .transpose()
            .map_err(|error| {
                ReadError::Decode(format!("invalid Keynote transition delay: {error}"))
            })?,
    );
    settings.set_autoplay_build_delay(
        raw.autoplay_build_delay()
            .map(Seconds::new)
            .transpose()
            .map_err(|error| ReadError::Decode(format!("invalid Keynote build delay: {error}")))?,
    );
    settings.set_idle_timer_active(raw.idle_timer_active());
    settings.set_idle_timer_delay(
        raw.idle_timer_delay()
            .map(Seconds::new)
            .transpose()
            .map_err(|error| ReadError::Decode(format!("invalid Keynote idle delay: {error}")))?,
    );
    settings.set_automatically_plays_upon_open(raw.automatically_plays_upon_open());
    settings
        .validate()
        .map_err(|error| ReadError::Decode(format!("invalid Keynote show settings: {error}")))?;
    Ok(settings)
}

fn property(dictionary: &plist::Dictionary, key: &str) -> Option<String> {
    dictionary
        .get(key)
        .and_then(plist::Value::as_string)
        .map(str::to_owned)
}

#[cfg(test)]
mod tests {
    use std::io::{Cursor, Write};
    use std::sync::Barrier;
    use std::thread;

    use super::*;
    use tempfile::NamedTempFile;

    fn assert_send_sync<T: Send + Sync>() {}

    fn native_fixture_path() -> std::path::PathBuf {
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/basic.key")
    }

    fn assert_input_limit(error: &ReadError, observed: u64, maximum: u64) {
        assert!(matches!(
            error,
            ReadError::Archive(litchi_iwa_archive::Error::Limit {
                kind: litchi_iwa_archive::LimitKind::InputBytes,
                observed: actual_observed,
                maximum: actual_maximum,
            }) if *actual_observed == observed && *actual_maximum == maximum
        ));
    }

    #[test]
    fn package_handles_are_send_sync() {
        assert_send_sync::<Package>();
    }

    #[test]
    fn focused_show_settings_does_not_initialize_full_slide_semantics()
    -> Result<(), Box<dyn std::error::Error>> {
        let package = Package::open(native_fixture_path())?;
        assert!(package.state.semantic.get().is_none());
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );

        let focused = package.show_settings()?;
        assert!(package.state.semantic.get().is_none());
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            0
        );

        assert_eq!(focused, *package.show()?.settings());
        assert!(package.state.semantic.get().is_some());
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            1
        );
        Ok(())
    }

    #[cfg(feature = "internal-iwork-source")]
    #[test]
    fn prepared_source_skips_redundant_keynote_classification()
    -> Result<(), Box<dyn std::error::Error>> {
        let source: Arc<[u8]> = std::fs::read(native_fixture_path())?.into();
        let Some(prepared) = PreparedSource::from_shared_bytes(Arc::clone(&source))? else {
            panic!("the native fixture must be recognized as Keynote");
        };

        let prepared_package =
            Package::__from_prepared_source(prepared, SemanticLimits::default())?;
        assert_eq!(prepared_package.state.source_classification_attempts, 0);

        let direct_package = Package::from_bytes(&source)?;
        assert_eq!(direct_package.state.source_classification_attempts, 1);
        Ok(())
    }

    #[test]
    fn semantic_initialization_is_single_flight_for_concurrent_readers()
    -> Result<(), Box<dyn std::error::Error>> {
        const WORKERS: usize = 16;

        let package = Package::open(native_fixture_path())?;
        let barrier = Arc::new(Barrier::new(WORKERS));
        let pointers = thread::scope(|scope| {
            let mut handles = Vec::with_capacity(WORKERS);
            for _ in 0..WORKERS {
                let package = package.snapshot();
                let barrier = Arc::clone(&barrier);
                handles.push(scope.spawn(move || {
                    barrier.wait();
                    package
                        .slides()
                        .map(|slides| slides.as_ptr() as usize)
                        .map_err(|error| error.to_string())
                }));
            }

            let mut pointers = Vec::with_capacity(WORKERS);
            for handle in handles {
                let result = handle
                    .join()
                    .map_err(|_panic| "concurrent semantic reader panicked".to_owned())?;
                pointers.push(result?);
            }
            Ok::<_, String>(pointers)
        })
        .map_err(std::io::Error::other)?;

        assert!(pointers.windows(2).all(|pair| pair[0] == pair[1]));
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            1
        );
        assert!(package.state.semantic.get().is_some());
        Ok(())
    }

    #[test]
    fn failed_semantic_initialization_remains_retryable() -> Result<(), Box<dyn std::error::Error>>
    {
        let bytes = std::fs::read(native_fixture_path())?;
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            1,
        )?;
        let package = Package::from_bytes_with_options(
            &bytes,
            ReadOptions::new(Limits::default(), semantic),
        )?;

        for _ in 0..2 {
            let Err(error) = package.show() else {
                panic!("the native fixture must exceed a one-byte semantic text limit");
            };
            assert!(matches!(
                error,
                ReadError::SemanticLimit {
                    kind: SemanticLimitKind::TextBytes,
                    maximum: 1,
                    ..
                }
            ));
            assert!(package.state.semantic.get().is_none());
        }
        assert_eq!(
            package
                .state
                .semantic_decode_attempts
                .load(Ordering::Relaxed),
            2
        );
        Ok(())
    }

    #[test]
    fn borrowed_input_exceeding_limit_is_rejected_before_copy()
    -> Result<(), Box<dyn std::error::Error>> {
        let limits = Limits::new(1, 1, 1, 1, 1)?;
        let Err(error) = Package::from_bytes_with_limits(&[0, 1], limits) else {
            panic!("oversized borrowed input should fail");
        };

        assert_input_limit(&error, 2, 1);
        Ok(())
    }

    #[test]
    fn path_input_exceeding_limit_is_rejected_before_materialization()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut file = NamedTempFile::new()?;
        file.write_all(&[0, 1])?;

        let limits = Limits::new(1, 1, 1, 1, 1)?;
        let Err(error) = Package::open_with_limits(file.path(), limits) else {
            panic!("oversized path input should fail");
        };

        assert_input_limit(&error, 2, 1);
        Ok(())
    }

    #[test]
    fn path_growth_past_limit_is_rejected_with_a_typed_limit()
    -> Result<(), Box<dyn std::error::Error>> {
        let limits = Limits::new(1, 1, 1, 1, 1)?;
        let mut reader = Cursor::new([0_u8; 64]);
        let Err(error) = read_source_with_reported_length(&mut reader, 0, limits) else {
            panic!("input growing beyond its reported length should fail");
        };

        assert_input_limit(&error, 2, 1);
        assert_eq!(reader.position(), 2);
        Ok(())
    }

    #[test]
    fn root_projection_keeps_the_base_archive_opaque() -> Result<(), Box<dyn std::error::Error>> {
        const OPAQUE_BYTES: usize = 256 * 1024;

        let mut source = vec![0x12, 0x02, 0x08, 0x2a, 0x1a];
        litchi_iwa_common::encode_varint_into(&mut source, u64::try_from(OPAQUE_BYTES)?);
        source.resize(source.len() + OPAQUE_BYTES, 0xff);

        assert_eq!(
            decode_root_show_identifier(&source, WireLimits::default())?,
            42
        );
        Ok(())
    }

    #[test]
    fn root_projection_requires_the_nested_show_identifier() {
        let Err(error) =
            decode_root_show_identifier(&[0x12, 0x00, 0x1a, 0x00], WireLimits::default())
        else {
            panic!("a show reference without its required identifier must fail");
        };
        assert!(matches!(error, ReadError::InvalidFormat(_)));
    }

    #[test]
    fn root_projection_forces_complete_reference_validation() {
        let Err(error) = decode_root_show_identifier(
            &[0x12, 0x04, 0x08, 0x2a, 0x12, 0x00, 0x1a, 0x00],
            WireLimits::default(),
        ) else {
            panic!("a known Reference field with the wrong wire type must fail");
        };
        assert!(matches!(error, ReadError::InvalidFormat(_)));
    }

    #[test]
    fn root_preflight_rejects_duplicate_show_identifiers() {
        let Err(error) = decode_root_show_identifier(
            &[0x12, 0x04, 0x08, 0x01, 0x08, 0x02, 0x1a, 0x00],
            WireLimits::default(),
        ) else {
            panic!("duplicate required identifiers must fail strict preflight");
        };
        assert!(matches!(error, ReadError::InvalidFormat(_)));
    }

    #[test]
    fn wire_limit_errors_preserve_counts_and_semantic_paths() {
        let path = SemanticPath::SlideBody { index: 3 };
        let common = litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Nesting,
            observed: 5,
            limit: 4,
        };
        assert!(matches!(
            map_wire_preflight_error(common.clone(), "test", path),
            ReadError::PayloadLimit {
                kind: PayloadLimitKind::Nesting,
                observed: 5,
                maximum: 4,
                path: SemanticPath::SlideBody { index: 3 },
            }
        ));

        let budget = SemanticBudget::new(SemanticLimits::default());
        assert!(matches!(
            budget.map_text_error(
                &TextWireError::Common(common),
                TextWireLimits::default(),
                path,
            ),
            ReadError::PayloadLimit {
                kind: PayloadLimitKind::Nesting,
                observed: 5,
                maximum: 4,
                path: SemanticPath::SlideBody { index: 3 },
            }
        ));
    }
}
