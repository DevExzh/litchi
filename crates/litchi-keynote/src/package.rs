//! Native Keynote package ingress and semantic decoding.
//!
//! This adapter owns the `.key` package boundary while preserving every raw
//! package member in its original byte stream. The archive, Snappy, detection,
//! and protobuf layers remain in their focused IWA infrastructure crates.

use std::fs::File;
use std::io::{Cursor, Read};
use std::path::Path;
use std::sync::{Arc, OnceLock};

use litchi_iwa_archive::ComponentCatalog;
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_detect::Format;
use litchi_iwa_protos::{kn, tswp};
use litchi_iwa_text::storage::Storage;
use prost::Message;
use thiserror::Error;

use crate::{
    AnimationType, Build, Document, Effect, Mode, Seconds, Settings, Show, Size, Slide, Transition,
};

/// Checked physical resource limits for Keynote package ingress.
pub use litchi_iwa_archive::Limits;

/// A result returned by a native Keynote package operation.
type ReadResult<T> = Result<T, ReadError>;

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
    /// The package properties plist could not be read.
    #[error("could not read Keynote package properties: {0}")]
    Metadata(#[from] plist::Error),
}

/// Cheaply cloneable parsed Keynote package with a lazy semantic snapshot.
///
/// The original package bytes remain available through [`Self::source_bytes`],
/// so unsupported IWA members and unmodeled protobuf fields are retained even
/// when callers inspect only the semantic presentation values.
#[derive(Debug, Clone)]
pub struct Package {
    state: Arc<State>,
}

#[derive(Debug)]
struct State {
    source: Arc<[u8]>,
    limits: Limits,
    components: ComponentCatalog,
    semantic: OnceLock<Document>,
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
        Self::open_with_limits(path, Limits::default())
    }

    /// Open a Keynote package from a filesystem path with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open`] while enforcing `limits`
    /// before materializing the complete source and at the physical ZIP and
    /// IWA boundaries.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> ReadResult<Self> {
        let source = read_source(path.as_ref(), limits)?;
        Self::from_source(source, limits)
    }

    /// Parse a Keynote package from complete ZIP bytes with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is not a bounded valid Keynote package.
    pub fn from_bytes(bytes: &[u8]) -> ReadResult<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Parse a Keynote package from complete ZIP bytes with explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error when the ZIP, Snappy/IWA components, document root,
    /// or requested resource profile is invalid.
    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> ReadResult<Self> {
        check_input_size(
            u64::try_from(bytes.len()).map_err(|_error| {
                ReadError::InvalidFormat("Keynote input length does not fit u64".to_owned())
            })?,
            limits,
        )?;
        let source = copy_source(bytes)?;
        Self::from_source(source, limits)
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

    fn from_source(source: Arc<[u8]>, limits: Limits) -> ReadResult<Self> {
        let components = ComponentCatalog::from_bytes_with_limits(source.as_ref(), limits)?;
        match litchi_iwa_detect::bytes(source.as_ref())? {
            Some(Format::Keynote) => {},
            Some(_) => return Err(ReadError::NotKeynote),
            None => {
                return Err(ReadError::InvalidFormat(
                    "package has no recognized iWork application root".to_owned(),
                ));
            },
        }

        let package = Self {
            state: Arc::new(State {
                source,
                limits,
                components,
                semantic: OnceLock::new(),
            }),
        };
        package.root_document()?;
        Ok(package)
    }

    /// Capture another handle to the same immutable parsed package.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow the original package bytes without normalizing unknown content.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.state.source
    }

    /// Return the checked physical limits used when this package was parsed.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.limits
    }

    /// Return the count of parsed native IWA components.
    #[must_use]
    pub fn component_count(&self) -> usize {
        self.state.components.len()
    }

    /// Iterate normalized names for all parsed native IWA components.
    pub fn component_names(&self) -> impl Iterator<Item = &str> {
        self.state
            .components
            .iter()
            .map(|component| component.name())
    }

    /// Extract textual content from native Keynote text storages in package order.
    ///
    /// Unsupported native objects remain preserved in [`Self::source_bytes`]
    /// and are skipped here rather than being interpreted speculatively.
    ///
    /// # Errors
    ///
    /// Returns an error only when the parsed package cannot maintain its
    /// validated semantic state.
    pub fn text(&self) -> ReadResult<String> {
        let mut parts = Vec::new();
        for object in self.objects() {
            for message in &object.messages {
                if let Ok(storage) = tswp::StorageArchive::decode(message.data.as_slice()) {
                    let text = storage.text.concat();
                    if !text.is_empty() {
                        parts.push(text);
                    }
                }
            }
        }
        Ok(parts.join("\n"))
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

        let catalog = Catalog::from_shared_bytes_with_limits(
            Arc::clone(&self.state.source),
            self.state.limits,
        )?;
        if let Some(properties) = catalog.iter().find(|entry| {
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
        self.root_document()?;
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
            total_objects: self.objects().count(),
            slide_count: self.slides()?.len(),
        })
    }

    fn semantic_document(&self) -> ReadResult<&Document> {
        if let Some(document) = self.state.semantic.get() {
            return Ok(document);
        }
        let document = Document::from_show(self.decode_show()?);
        let _ = self.state.semantic.set(document);
        self.state
            .semantic
            .get()
            .ok_or_else(|| ReadError::Decode("semantic snapshot was not initialized".to_owned()))
    }

    fn root_document(&self) -> ReadResult<kn::DocumentArchive> {
        let mut roots = self
            .state
            .components
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
        object
            .messages
            .iter()
            .find_map(|message| kn::DocumentArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                ReadError::InvalidFormat("missing Keynote root document payload".to_owned())
            })
    }

    fn decode_show(&self) -> ReadResult<Show> {
        let document = self.root_document()?;
        let show_object = self.object(document.show.identifier).ok_or_else(|| {
            ReadError::Decode(format!(
                "Keynote show object {} is missing",
                document.show.identifier
            ))
        })?;
        let show = show_object
            .messages
            .iter()
            .find_map(|message| kn::ShowArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| ReadError::Decode("Keynote show payload is missing".to_owned()))?;

        let mut builder = Show::builder();
        builder.set_settings(settings_from_show(&show)?);
        builder.set_title(self.object_text(show_object).into_iter().next());
        for (index, slide_id) in self.slide_ids(&show)?.into_iter().enumerate() {
            let object = self.object(slide_id).ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote slide tree references missing object {slide_id}"
                ))
            })?;
            builder.push_slide(self.parse_slide(index, object)?);
        }
        Ok(builder.build())
    }

    fn slide_ids(&self, show: &kn::ShowArchive) -> ReadResult<Vec<u64>> {
        let mut ids = Vec::with_capacity(show.slide_tree.slides.len());
        for reference in &show.slide_tree.slides {
            let node = self.object(reference.identifier).ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote slide node {} is missing",
                    reference.identifier
                ))
            })?;
            let node = node
                .messages
                .iter()
                .find_map(|message| kn::SlideNodeArchive::decode(message.data.as_slice()).ok())
                .ok_or_else(|| {
                    ReadError::Decode(format!(
                        "object {} has no Keynote slide-node payload",
                        reference.identifier
                    ))
                })?;
            if let Some(slide) = node.slide {
                ids.push(slide.identifier);
            }
        }
        Ok(ids)
    }

    fn parse_slide(&self, index: usize, object: &ArchiveObject) -> ReadResult<Slide> {
        let mut builder = Slide::builder(index);
        let text = self.object_text(object);
        if let Some(title) = text.first() {
            builder.set_title(Some(title.clone()));
        }
        for value in text.into_iter().skip(1) {
            builder.push_text(value);
        }

        if let Some(slide) = object
            .messages
            .iter()
            .find_map(|message| kn::SlideArchive::decode(message.data.as_slice()).ok())
        {
            if let Some(name) = slide.name.filter(|name| !name.is_empty()) {
                builder.set_title(Some(name));
            }
            for build in &slide.builds {
                builder.push_build(self.extract_build(build.identifier)?);
            }
            builder.set_transition(Some(self.transition(&slide.transition)?));

            let title = slide
                .title_placeholder
                .as_ref()
                .map(|reference| reference.identifier);
            let body = slide
                .body_placeholder
                .as_ref()
                .map(|reference| reference.identifier);
            if let Some(identifier) = title {
                let text = self.drawable_text(identifier)?;
                if !text.is_empty() {
                    builder.set_title(Some(text));
                }
            }
            if let Some(identifier) = body {
                let text = self.drawable_text(identifier)?;
                if !text.is_empty() {
                    builder.push_text(text);
                }
            }
            for drawable in &slide.owned_drawables {
                if Some(drawable.identifier) == title || Some(drawable.identifier) == body {
                    continue;
                }
                let text = self.drawable_text(drawable.identifier)?;
                if !text.is_empty() {
                    builder.push_text(text);
                }
            }
            if let Some(note) = slide.note {
                let text = self.notes_text(note.identifier)?;
                if !text.is_empty() {
                    builder.set_notes(Some(text));
                }
            }
        }

        if let Some(storage) = self.text_storage(object) {
            builder.push_text_storage(storage);
        }
        Ok(builder.build())
    }

    fn extract_build(&self, identifier: u64) -> ReadResult<Build> {
        let object = self.object(identifier).ok_or_else(|| {
            ReadError::Decode(format!("Keynote build object {identifier} is missing"))
        })?;
        let build = object
            .messages
            .iter()
            .find_map(|message| kn::BuildArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote build object {identifier} has no build payload"
                ))
            })?;
        let animation = AnimationType::from_identifier(&build.delivery).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote build identifier: {error}"))
        })?;
        #[allow(deprecated)]
        let duration = build
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or(build.attributes.database_duration)
            .or(build.duration)
            .unwrap_or(0.0);
        let duration = Seconds::new(duration).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote build duration: {error}"))
        })?;
        Ok(Build::new(animation, duration))
    }

    fn transition(&self, transition: &kn::TransitionArchive) -> ReadResult<Transition> {
        #[allow(deprecated)]
        let duration = transition
            .attributes
            .animation_attributes
            .as_ref()
            .and_then(|attributes| attributes.duration)
            .or(transition.attributes.database_duration)
            .unwrap_or(0.0);
        let duration = Seconds::new(duration).map_err(|error| {
            ReadError::Decode(format!("invalid Keynote transition duration: {error}"))
        })?;
        #[allow(deprecated)]
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

    fn drawable_text(&self, identifier: u64) -> ReadResult<String> {
        let drawable = self.object(identifier).ok_or_else(|| {
            ReadError::Decode(format!("Keynote drawable object {identifier} is missing"))
        })?;
        let storage = drawable.messages.iter().find_map(|message| {
            kn::PlaceholderArchive::decode(message.data.as_slice())
                .ok()
                .and_then(|placeholder| placeholder.super_.owned_storage)
                .or_else(|| {
                    tswp::ShapeInfoArchive::decode(message.data.as_slice())
                        .ok()
                        .and_then(|shape| shape.owned_storage)
                })
        });
        let Some(storage) = storage else {
            return Ok(String::new());
        };
        let storage_id = storage.identifier;
        let storage = self.object(storage_id).ok_or_else(|| {
            ReadError::Decode(format!(
                "Keynote drawable storage object {} is missing",
                storage_id
            ))
        })?;
        self.text_storage(storage)
            .map(Storage::into_text)
            .ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote drawable storage object {} has no text payload",
                    storage_id
                ))
            })
    }

    fn notes_text(&self, identifier: u64) -> ReadResult<String> {
        let note = self.object(identifier).ok_or_else(|| {
            ReadError::Decode(format!("Keynote notes object {identifier} is missing"))
        })?;
        let note = note
            .messages
            .iter()
            .find_map(|message| kn::NoteArchive::decode(message.data.as_slice()).ok())
            .ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote notes object {identifier} has no note payload"
                ))
            })?;
        let storage = self
            .object(note.contained_storage.identifier)
            .ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote notes storage object {} is missing",
                    note.contained_storage.identifier
                ))
            })?;
        self.text_storage(storage)
            .map(Storage::into_text)
            .ok_or_else(|| {
                ReadError::Decode(format!(
                    "Keynote notes storage object {} has no text payload",
                    note.contained_storage.identifier
                ))
            })
    }

    fn objects(&self) -> impl Iterator<Item = &ArchiveObject> {
        self.state
            .components
            .iter()
            .flat_map(|component| component.archive().objects.iter())
    }

    fn object(&self, identifier: u64) -> Option<&ArchiveObject> {
        self.state
            .components
            .iter()
            .find_map(|component| component.archive().object(identifier))
    }

    fn object_text(&self, object: &ArchiveObject) -> Vec<String> {
        object
            .messages
            .iter()
            .filter_map(|message| tswp::StorageArchive::decode(message.data.as_slice()).ok())
            .map(|storage| storage.text.concat())
            .filter(|text| !text.is_empty())
            .collect()
    }

    fn text_storage(&self, object: &ArchiveObject) -> Option<Storage> {
        object
            .messages
            .iter()
            .find_map(|message| tswp::StorageArchive::decode(message.data.as_slice()).ok())
            .map(|storage| Storage::from_text(storage.text.concat()))
    }
}

fn read_source(path: &Path, limits: Limits) -> ReadResult<Arc<[u8]>> {
    let mut file = File::open(path)?;
    let length = file.metadata()?.len();
    check_input_size(length, limits)?;

    let maximum = usize::try_from(limits.max_input_bytes()).map_err(|_error| {
        ReadError::InvalidFormat("Keynote input limit does not fit usize".to_owned())
    })?;
    let capacity = usize::try_from(length).map_err(|_error| {
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
            if file.read(&mut extra)? != 0 {
                return Err(input_limit_error(
                    limits.max_input_bytes().saturating_add(1),
                    limits,
                ));
            }
            break;
        }

        let read_limit = remaining.min(buffer.len());
        let read = file.read(&mut buffer[..read_limit])?;
        if read == 0 {
            break;
        }
        bytes.try_reserve(read).map_err(|_error| {
            ReadError::Archive(litchi_iwa_archive::Error::Allocation {
                resource: "Keynote package input",
                amount: read,
            })
        })?;
        bytes.extend_from_slice(&buffer[..read]);
    }

    Ok(bytes.into())
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

fn settings_from_show(show: &kn::ShowArchive) -> ReadResult<Settings> {
    let size = Size::new(show.size.width, show.size.height)
        .map_err(|error| ReadError::Decode(format!("invalid Keynote show size: {error}")))?;
    let mut settings = Settings::new(size);
    settings.set_slide_numbers_visible(show.slide_numbers_visible);
    settings.set_loop_presentation(show.loop_presentation);
    settings
        .set_mode(show.mode.map(Mode::from_raw))
        .map_err(|error| ReadError::Decode(format!("invalid Keynote show mode: {error}")))?;
    settings.set_autoplay_transition_delay(
        show.autoplay_transition_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| {
                ReadError::Decode(format!("invalid Keynote transition delay: {error}"))
            })?,
    );
    settings.set_autoplay_build_delay(
        show.autoplay_build_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| ReadError::Decode(format!("invalid Keynote build delay: {error}")))?,
    );
    settings.set_idle_timer_active(show.idle_timer_active);
    settings.set_idle_timer_delay(
        show.idle_timer_delay
            .map(Seconds::new)
            .transpose()
            .map_err(|error| ReadError::Decode(format!("invalid Keynote idle delay: {error}")))?,
    );
    settings.set_automatically_plays_upon_open(show.automatically_plays_upon_open);
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
    use std::io::Write;

    use super::*;
    use tempfile::NamedTempFile;

    fn assert_send_sync<T: Send + Sync>() {}

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
}
