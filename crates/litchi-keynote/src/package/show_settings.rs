//! Immutable transactions for Keynote presentation settings.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_document_codec;
use thiserror::Error;

use super::{
    DOCUMENT_MESSAGE_TYPE, Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE, SemanticBudget,
    SemanticPath, decode_show_settings_snapshot, preflight_document, preflight_show,
    settings_from_show_projection,
};
use crate::Seconds;
use crate::show::{Settings, Size};

const SHOW_SIZE_FIELD: u32 = 4;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;
const SLIDE_NUMBERS_VISIBLE_FIELD: u32 = 6;
const LOOP_PRESENTATION_FIELD: u32 = 8;
const MODE_FIELD: u32 = 9;
const AUTOPLAY_TRANSITION_DELAY_FIELD: u32 = 10;
const AUTOPLAY_BUILD_DELAY_FIELD: u32 = 11;
const IDLE_TIMER_ACTIVE_FIELD: u32 = 15;
const IDLE_TIMER_DELAY_FIELD: u32 = 16;
const AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD: u32 = 18;
const SCALAR_FIELDS: [u32; 8] = [
    SLIDE_NUMBERS_VISIBLE_FIELD,
    LOOP_PRESENTATION_FIELD,
    MODE_FIELD,
    AUTOPLAY_TRANSITION_DELAY_FIELD,
    AUTOPLAY_BUILD_DELAY_FIELD,
    IDLE_TIMER_ACTIVE_FIELD,
    IDLE_TIMER_DELAY_FIELD,
    AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD,
];

/// A finite resource governed while show settings are prepared or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member or IWA value.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic rich-text fragments.
    TextFragments,
    /// Aggregate retained semantic text.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while staging or publishing show settings.
///
/// Values contain only semantic failure categories and bounded resource
/// measurements. Package member names, object identifiers, raw bytes, and
/// lower-layer diagnostic strings remain private.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// The package has no physical source that can back an exact patch.
    #[error("this Keynote source does not support physical show-settings edits")]
    UnsupportedSource,
    /// The source package or selected show payload is structurally invalid.
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A finite retained resource ceiling was exceeded.
    #[error("Keynote show-settings {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote show-settings transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full semantic readback did not reproduce the requested settings.
    #[error("the edited Keynote show settings failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote show-settings patch does not match the exact source package")]
    PatchConflict,
}

/// A mutable settings snapshot staged against one immutable Keynote package.
#[derive(Debug)]
pub struct Edit<'a> {
    source: &'a Package,
    before: Settings,
    settings: Settings,
}

impl<'a> Edit<'a> {
    fn new(source: &'a Package) -> Result<Self, Error> {
        physical_source_catalog(source)?;
        let before = strict_package_settings(source)?;
        Ok(Self {
            source,
            before,
            settings: before,
        })
    }

    /// Return the settings that would be published by this edit.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the staged semantic settings.
    ///
    /// [`Settings`] is valid by construction.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// An exact semantic no-op reuses the package source allocation. A change
    /// is published only after full source and candidate validation plus a
    /// strict Buffa semantic readback under the retained limits.
    ///
    /// # Costs
    ///
    /// A no-op hashes the retained artifact but does not inspect caches,
    /// reassemble, or reopen it. A change rewrites one IWA component, may
    /// delete zero to three root previews for size or slide-number changes,
    /// and fully reopens and verifies the candidate. Playback-only changes
    /// preserve root previews exactly.
    ///
    /// # Errors
    ///
    /// Returns an error without modifying `source` when its physical
    /// provenance, semantic graph, wire payload, resource profile, allocation,
    /// or readback invariant is rejected.
    pub fn commit(self) -> Result<Commit, Error> {
        let source_catalog = physical_source_catalog(self.source)?;
        let source_bytes = source_catalog.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);

        if self.before == self.settings {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    before: self.before,
                    after: self.settings,
                    source_preview_count: 0,
                    target_preview_count: 0,
                    touched_components: 0,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }

        if !source_catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let RewriteResult {
            package,
            touched_components,
            source_preview_count,
            target_preview_count,
        } = rewrite_package_settings(self.source, self.before, self.settings)?;
        let target_bytes = physical_source_catalog(&package)?.shared_source();
        let target_fingerprint = fingerprint(&target_bytes);
        let deleted_previews = source_preview_count.saturating_sub(target_preview_count);
        Ok(Commit {
            package,
            patch: Patch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                before: self.before,
                after: self.settings,
                source_preview_count,
                target_preview_count,
                touched_components,
            },
            diagnostics: Diagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// An exact-source-checked, reversible semantic show-settings patch.
///
/// Native identifiers, package member names, and exact source/target bytes are
/// retained privately. Fingerprints are compact diagnostics; exact private
/// byte comparison authorizes patch application.
///
/// A patch retains two complete package artifacts. Clone and inversion are
/// `O(1)` `Arc` operations; equality and exact-artifact authorization can read
/// `O(package bytes)`. It is process-local in-memory state, not a compact or
/// durable serialization format.
#[derive(Clone, PartialEq)]
pub struct Patch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Settings,
    after: Settings,
    source_preview_count: usize,
    target_preview_count: usize,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the semantic settings required from the patch source.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Return the semantic settings produced by the patch target.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether the patch preserves the semantic settings and exact bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source_preview_count == self.target_preview_count
            && (Arc::ptr_eq(&self.source_bytes, &self.target_bytes)
                || self.source_bytes.as_ref() == self.target_bytes.as_ref())
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            before: self.after,
            after: self.before,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            touched_components: self.touched_components,
        }
    }
}

/// Compact evidence describing one show-settings commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of root preview entries deleted by this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable show-settings transaction.
#[must_use = "a Keynote show-settings commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read validated presentation dimensions and playback settings.
    ///
    /// This focused reader runs the strict bounded show preflight and Buffa
    /// settings projection directly over the retained show payload. It does
    /// not initialize or retain the full semantic slide document.
    ///
    /// # Costs
    ///
    /// Scans the rooted Document and Show payloads under retained bounds.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the show payload cannot be projected safely
    /// under this package's retained physical and semantic limits.
    pub fn show_settings(&self) -> Result<Settings, Error> {
        strict_package_settings(self)
    }

    /// Start a focused edit of presentation dimensions and playback settings.
    ///
    /// The transaction's before value comes directly from the strict bounded
    /// Buffa show projection, not from a generated mutable protobuf model.
    ///
    /// # Costs
    ///
    /// Performs the same focused rooted read as [`Self::show_settings`].
    /// Staging a valid [`Settings`] replacement is `O(1)`.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the package has no physical patch source or
    /// its show payload cannot be projected safely under the retained limits.
    pub fn edit_show_settings(&self) -> Result<Edit<'_>, Error> {
        Edit::new(self)
    }

    /// Apply an exact-source-checked show-settings patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    ///
    /// # Costs
    ///
    /// Exact-source authorization can compare `O(package bytes)`. A no-op
    /// reuses the source without cache inspection, reassembly, or reopening;
    /// a changed patch reopens and verifies its retained target.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`, or another typed validation
    /// error when its retained target cannot be published safely.
    pub fn apply_show_settings(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_catalog = physical_source_catalog(self)?;
        let source_bytes = source_catalog.shared_source();
        let exact_source = Arc::ptr_eq(&source_bytes, &patch.source_bytes)
            || (fingerprint(source_catalog.source_bytes()) == patch.source_fingerprint
                && source_catalog.source_bytes() == patch.source_bytes.as_ref());
        if !exact_source {
            return Err(Error::PatchConflict);
        }

        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }

        if root_preview_count(source_catalog)? != patch.source_preview_count {
            return Err(Error::PatchConflict);
        }

        let admitted = select_rooted_show(self, true, true)?;
        let admitted_settings = admitted
            .payload
            .map(|payload| strict_payload_settings(self, payload))
            .transpose()?
            .unwrap_or_default();
        if admitted_settings != patch.before {
            return Err(Error::PatchConflict);
        }

        if !source_catalog.source_is_exact()
            || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target_bytes), self.state.options)
                .map_err(map_read_error)?;
        verify_candidate(
            self,
            &candidate,
            patch.before,
            patch.after,
            patch.source_preview_count,
            patch.target_preview_count,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(
                patch.touched_components,
                patch
                    .source_preview_count
                    .saturating_sub(patch.target_preview_count),
            ),
        })
    }
}

#[derive(Debug, Clone, Copy)]
enum ScalarValue {
    Varint(u64),
    Fixed64(u64),
}

impl ScalarValue {
    const fn wire_type(self) -> u8 {
        match self {
            Self::Varint(_) => 0,
            Self::Fixed64(_) => 1,
        }
    }

    fn encoded_len(self, field_number: u32) -> usize {
        let key = (u64::from(field_number) << 3) | u64::from(self.wire_type());
        encoded_len(key)
            + match self {
                Self::Varint(value) => encoded_len(value),
                Self::Fixed64(_) => 8,
            }
    }

    fn append(self, output: &mut Vec<u8>, field_number: u32) {
        let key = (u64::from(field_number) << 3) | u64::from(self.wire_type());
        encode_varint_into(output, key);
        match self {
            Self::Varint(value) => encode_varint_into(output, value),
            Self::Fixed64(value) => output.extend_from_slice(&value.to_le_bytes()),
        }
    }
}

struct RewriteResult {
    package: Package,
    touched_components: usize,
    source_preview_count: usize,
    target_preview_count: usize,
}

struct ShowSelection<'source> {
    identifier: u64,
    root_component_name: &'source str,
    component_name: Option<&'source str>,
    payload: Option<&'source [u8]>,
}

fn strict_package_settings(package: &Package) -> Result<Settings, Error> {
    let selected = select_rooted_show(package, false, false)?;
    match selected.payload {
        Some(payload) => strict_payload_settings(package, payload),
        None if selected.identifier == 0 => Ok(Settings::default()),
        None => Err(Error::InvalidSource),
    }
}

fn select_rooted_show(
    package: &Package,
    mutation_guards: bool,
    canonical_framing: bool,
) -> Result<ShowSelection<'_>, Error> {
    let catalog = physical_source_catalog(package)?;
    let mut root_components = catalog
        .components()
        .iter()
        .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
    let root_component = root_components.next().ok_or(Error::InvalidSource)?;
    if root_components.next().is_some() {
        return Err(Error::InvalidSource);
    }
    if canonical_framing {
        validate_component_framing(package, root_component.name(), root_component.archive())?;
    }
    let root = root_component
        .archive()
        .object(1)
        .ok_or(Error::InvalidSource)?;
    let (root_message_index, root_payload) = selected_message(root, DOCUMENT_MESSAGE_TYPE)?;
    if mutation_guards {
        validate_selected_metadata(root, root_message_index)?;
    }

    let wire_limits = package.semantic_wire_limits().map_err(map_read_error)?;
    preflight_document(root_payload, wire_limits).map_err(map_read_error)?;
    let recursion_limit =
        u32::try_from(wire_limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    let reference = keynote_document_codec::decode_show_reference(
        root_payload,
        keynote_document_codec::DecodeOptions::new(root_payload.len(), recursion_limit)
            .with_max_fields(wire_limits.max_fields())
            .with_max_work_bytes(wire_limits.max_rewrite_work()),
    )
    .map_err(|_error| Error::InvalidSource)?;
    if reference.deprecated_is_external() == Some(true) {
        return Err(Error::InvalidSource);
    }
    let show_identifier = reference.identifier();
    validate_root_show_metadata(root, root_message_index, show_identifier)?;
    if show_identifier == 0 {
        return Ok(ShowSelection {
            identifier: 0,
            root_component_name: root_component.name(),
            component_name: None,
            payload: None,
        });
    }

    let mut show_components = catalog
        .components()
        .iter()
        .filter(|component| component.archive().object(show_identifier).is_some());
    let show_component = show_components.next().ok_or(Error::InvalidSource)?;
    if show_components.next().is_some() {
        return Err(Error::InvalidSource);
    }
    if show_component.name() != root_component.name() && canonical_framing {
        validate_component_framing(package, show_component.name(), show_component.archive())?;
    }
    let show = show_component
        .archive()
        .object(show_identifier)
        .ok_or(Error::InvalidSource)?;
    let (show_message_index, show_payload) = selected_message(show, SHOW_MESSAGE_TYPE)?;
    if mutation_guards {
        validate_selected_metadata(show, show_message_index)?;
    }
    Ok(ShowSelection {
        identifier: show_identifier,
        root_component_name: root_component.name(),
        component_name: Some(show_component.name()),
        payload: Some(show_payload),
    })
}

fn selected_message(object: &ArchiveObject, message_type: u32) -> Result<(usize, &[u8]), Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn validate_selected_metadata(object: &ArchiveObject, message_index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_root_show_metadata(
    root: &ArchiveObject,
    message_index: usize,
    show_identifier: u64,
) -> Result<(), Error> {
    let info = root
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    // The aggregate reference index may contain unrelated root references,
    // but must record the selected non-null show exactly once. A field-local
    // `[2]` index, when present, is the precise ownership proof for that
    // selected show reference.
    let aggregate_occurrences = info
        .object_references
        .iter()
        .filter(|identifier| **identifier == show_identifier)
        .count();
    if (show_identifier == 0 && aggregate_occurrences != 0)
        || (show_identifier != 0 && aggregate_occurrences != 1)
    {
        return Err(Error::InvalidSource);
    }
    let mut field_path_seen = false;
    for field in &info.field_infos {
        if field.path.as_slice() == [2] {
            if show_identifier == 0 {
                if !field.object_references.is_empty() {
                    return Err(Error::InvalidSource);
                }
                continue;
            }
            if field_path_seen || field.object_references.as_slice() != [show_identifier] {
                return Err(Error::InvalidSource);
            }
            field_path_seen = true;
        } else if show_identifier != 0 && field.object_references.contains(&show_identifier) {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn validate_component_framing(
    package: &Package,
    component_name: &str,
    archive: &Archive,
) -> Result<(), Error> {
    let catalog = physical_source_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let physical_limits = package.state.options.archive();
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), archive)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), Error> {
    for object in &archive.objects {
        let offset =
            usize::try_from(object.header_offset).map_err(|_error| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| Error::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_error| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(Error::InvalidSource)?
                != object.data_offset
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn strict_payload_settings(package: &Package, payload: &[u8]) -> Result<Settings, Error> {
    let wire_limits = package.semantic_wire_limits().map_err(map_read_error)?;
    let mut budget = SemanticBudget::new(package.semantic_limits());
    budget
        .charge_references(1, SemanticPath::Show)
        .map_err(map_read_error)?;
    preflight_show(payload, wire_limits, &mut budget).map_err(map_read_error)?;
    let snapshot =
        decode_show_settings_snapshot(payload, package.semantic_limits().max_slides(), wire_limits)
            .map_err(map_read_error)?;
    settings_from_show_projection(snapshot.size(), snapshot.raw_settings()).map_err(map_read_error)
}

fn rewrite_package_settings(
    source: &Package,
    before: Settings,
    after: Settings,
) -> Result<RewriteResult, Error> {
    let source_catalog = editable_source_catalog(source)?;
    let selected = select_rooted_show(source, true, false)?;
    if selected.identifier == 0 {
        return Err(Error::UnsupportedSource);
    }
    let show_identifier = selected.identifier;
    let component_name = selected.component_name.ok_or(Error::InvalidSource)?;
    if selected.root_component_name != component_name {
        let root_component = source_catalog
            .components()
            .get(selected.root_component_name)
            .ok_or(Error::InvalidSource)?;
        validate_component_framing(
            source,
            selected.root_component_name,
            root_component.archive(),
        )?;
    }
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }

    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    drop(stream);
    let object = archive
        .object(show_identifier)
        .ok_or(Error::InvalidSource)?;
    let (message_index, message_data) = selected_message(object, SHOW_MESSAGE_TYPE)?;
    validate_selected_metadata(object, message_index)?;
    if strict_payload_settings(source, message_data)? != before {
        return Err(Error::InvalidSource);
    }
    let rewritten = rewrite_show_payload(
        message_data,
        before,
        after,
        source.wire_limits().map_err(map_wire_error)?,
    )?;
    if strict_payload_settings(source, &rewritten)? != after {
        return Err(Error::Verification);
    }

    archive
        .object_mut(show_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: SHOW_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let rewritten_archive = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    drop(archive);
    let compressed = SnappyStream::compress(&rewritten_archive).map_err(map_core_error)?;
    drop(rewritten_archive);
    let preview_plan =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    let source_preview_count = preview_plan.len();
    let invalidate_rendering = rendering_changed(before, after);
    let deleted_previews = if invalidate_rendering {
        preview_plan.names()
    } else {
        &[]
    };
    let output = source_catalog
        .package()
        .reassemble_with_deletions_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            deleted_previews,
            physical_limits,
        )
        .map_err(map_archive_error)?;
    drop(compressed);
    drop(preview_plan);

    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    let target_preview_count = if invalidate_rendering {
        0
    } else {
        source_preview_count
    };
    verify_candidate(
        source,
        &candidate,
        before,
        after,
        source_preview_count,
        target_preview_count,
    )?;
    Ok(RewriteResult {
        package: candidate,
        touched_components: 1,
        source_preview_count,
        target_preview_count,
    })
}

fn rendering_changed(before: Settings, after: Settings) -> bool {
    before.size() != after.size() || before.slide_numbers_visible() != after.slide_numbers_visible()
}

fn root_preview_count(catalog: &SourceCatalog) -> Result<usize, Error> {
    super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map(|plan| plan.len())
        .map_err(map_rendering_invalidation_error)
}

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    before: Settings,
    after: Settings,
    expected_source_previews: usize,
    expected_target_previews: usize,
) -> Result<(), Error> {
    // The caller admitted the source framing before constructing or applying
    // the retained target. Reproject ownership here without decompressing the
    // same source components a second time; the freshly reopened candidate
    // still receives the complete framing guard below.
    let source_selection = select_rooted_show(source, true, false)?;
    let candidate_selection = select_rooted_show(candidate, true, true)?;
    if source_selection.identifier == 0
        || source_selection.identifier != candidate_selection.identifier
        || source_selection.component_name != candidate_selection.component_name
    {
        return Err(Error::Verification);
    }
    let source_settings = source_selection
        .payload
        .ok_or(Error::Verification)
        .and_then(|payload| strict_payload_settings(source, payload))?;
    let candidate_settings = candidate_selection
        .payload
        .ok_or(Error::Verification)
        .and_then(|payload| strict_payload_settings(candidate, payload))?;
    if source_settings != before || candidate_settings != after {
        return Err(Error::Verification);
    }

    let source_catalog = physical_source_catalog(source)?;
    let candidate_catalog = physical_source_catalog(candidate)?;
    let source_preview_plan =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    let candidate_preview_plan =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(map_rendering_invalidation_error)?;
    if source_preview_plan.len() != expected_source_previews
        || candidate_preview_plan.len() != expected_target_previews
    {
        return Err(Error::Verification);
    }
    if rendering_changed(before, after) {
        if expected_source_previews != expected_target_previews {
            if expected_source_previews.min(expected_target_previews) != 0 {
                return Err(Error::Verification);
            }
        } else if expected_target_previews != 0 {
            return Err(Error::Verification);
        }
        if expected_target_previews == 0
            && !super::rendering_invalidation::root_previews_absent(candidate_catalog.package())
                .map_err(map_rendering_invalidation_error)?
        {
            return Err(Error::Verification);
        }
    } else if expected_source_previews != expected_target_previews
        || !super::rendering_invalidation::root_previews_preserved(
            source_catalog.package(),
            candidate_catalog.package(),
        )
        .map_err(map_rendering_invalidation_error)?
    {
        return Err(Error::Verification);
    }

    let show_component_name = source_selection.component_name.ok_or(Error::Verification)?;
    verify_package_members(
        source_catalog,
        candidate_catalog,
        show_component_name,
        source_preview_plan.names(),
        candidate_preview_plan.names(),
    )?;
    verify_component_objects(
        source_catalog,
        candidate_catalog,
        show_component_name,
        source_selection.identifier,
    )
}

fn verify_package_members(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    show_component_name: &str,
    source_previews: &[&str],
    candidate_previews: &[&str],
) -> Result<(), Error> {
    let mut source_entries = source
        .package()
        .iter()
        .filter(|entry| !source_previews.contains(&entry.name()));
    let mut candidate_entries = candidate
        .package()
        .iter()
        .filter(|entry| !candidate_previews.contains(&entry.name()));
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(before), Some(after))
                if before.name() == after.name()
                    && if before.name() == show_component_name {
                        selected_package_member_preserved(before, after)
                    } else {
                        package_member_preserved(before, after)
                    } => {},
            (None, None) => return Ok(()),
            _ => return Err(Error::Verification),
        }
    }
}

fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && super::rendering_invalidation::central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

pub(super) fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_local_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;

    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    let Some(source_header_length) = zip_local_header_length(source_record) else {
        return false;
    };
    let Some(candidate_header_length) = zip_local_header_length(candidate_record) else {
        return false;
    };
    if source_header_length != candidate_header_length
        || source_record[..CRC_AND_SIZES.start] != candidate_record[..CRC_AND_SIZES.start]
        || source_record[CRC_AND_SIZES.end..source_header_length]
            != candidate_record[CRC_AND_SIZES.end..candidate_header_length]
    {
        return false;
    }

    let Some(source_payload_end) = source_header_length
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= source_record.len())
    else {
        return false;
    };
    let Some(candidate_payload_end) = candidate_header_length
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= candidate_record.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &source_record[source_payload_end..],
        &candidate_record[candidate_payload_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name_length = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra_length = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name_length)?
        .checked_add(extra_length)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let source_descriptor = if source.starts_with(b"PK\x07\x08") {
        4
    } else {
        0
    };
    let candidate_descriptor = if candidate.starts_with(b"PK\x07\x08") {
        4
    } else {
        0
    };
    source_descriptor == candidate_descriptor
        && source.len() == candidate.len()
        && source.len() >= source_descriptor + 12
        && source[..source_descriptor] == candidate[..candidate_descriptor]
        && source[source_descriptor + 12..] == candidate[candidate_descriptor + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;

    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
            == candidate[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn verify_component_objects(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    show_component_name: &str,
    show_identifier: u64,
) -> Result<(), Error> {
    let before_component = source
        .components()
        .get(show_component_name)
        .ok_or(Error::Verification)?;
    let after_component = candidate
        .components()
        .get(show_component_name)
        .ok_or(Error::Verification)?;
    if before_component.archive().objects.len() != after_component.archive().objects.len() {
        return Err(Error::Verification);
    }
    let mut selected_seen = false;
    for (before_object, after_object) in before_component
        .archive()
        .objects
        .iter()
        .zip(&after_component.archive().objects)
    {
        let before_identifier = before_object
            .archive_info
            .identifier
            .ok_or(Error::Verification)?;
        if after_object.archive_info.identifier != Some(before_identifier) {
            return Err(Error::Verification);
        }
        if before_identifier == show_identifier {
            if std::mem::replace(&mut selected_seen, true) {
                return Err(Error::Verification);
            }
            verify_selected_show_object(before_object, after_object)?;
        } else if before_object.archive_info != after_object.archive_info
            || before_object.messages != after_object.messages
        {
            return Err(Error::Verification);
        }
    }
    if selected_seen {
        Ok(())
    } else {
        Err(Error::Verification)
    }
}

fn verify_selected_show_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
) -> Result<(), Error> {
    let (source_index, _source_payload) = selected_message(source, SHOW_MESSAGE_TYPE)?;
    let (candidate_index, _candidate_payload) = selected_message(candidate, SHOW_MESSAGE_TYPE)?;
    if source_index != candidate_index
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
    {
        return Err(Error::Verification);
    }
    for (index, (before, after)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if before.type_ != after.type_ || (index != source_index && before != after) {
            return Err(Error::Verification);
        }
    }
    for (index, (before, after)) in source
        .archive_info
        .message_infos
        .iter()
        .zip(&candidate.archive_info.message_infos)
        .enumerate()
    {
        if index == source_index {
            if !message_info_preserved_except_length(before, after) {
                return Err(Error::Verification);
            }
        } else if before != after {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn message_info_preserved_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn rewrite_show_payload(
    source: &[u8],
    before: Settings,
    after: Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output_length = 0usize;
    let mut output_fields = view.len();
    let mut seen_scalars = 0u32;
    let mut saw_size = false;
    let mut size_payload = None;
    let mut nested_work = 0usize;

    for field in view.fields() {
        let replacement_length = if field.number() == SHOW_SIZE_FIELD {
            if saw_size || field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            saw_size = true;
            if before.size() == after.size() {
                field.raw().len()
            } else {
                nested_work = inspect_size_payload(field.payload(), limits)?;
                size_payload = Some(field.payload());
                field.raw().len()
            }
        } else if is_scalar_settings_field(field.number()) {
            let bit = scalar_presence_bit(field.number());
            if seen_scalars & bit != 0 || field.wire_type() != scalar_wire_type(field.number()) {
                return Err(Error::InvalidSource);
            }
            seen_scalars |= bit;
            if scalar_changed(field.number(), before, after) {
                if let Some(value) = scalar_value(field.number(), after) {
                    value.encoded_len(field.number())
                } else {
                    output_fields = output_fields.checked_sub(1).ok_or(Error::InvalidSource)?;
                    0
                }
            } else {
                field.raw().len()
            }
        } else {
            field.raw().len()
        };
        output_length = output_length
            .checked_add(replacement_length)
            .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
    }
    if !saw_size {
        return Err(Error::InvalidSource);
    }

    for field_number in SCALAR_FIELDS {
        let present = seen_scalars & scalar_presence_bit(field_number) != 0;
        if present != scalar_value(field_number, before).is_some() {
            return Err(Error::InvalidSource);
        }
        if !present
            && scalar_changed(field_number, before, after)
            && let Some(value) = scalar_value(field_number, after)
        {
            output_length = output_length
                .checked_add(value.encoded_len(field_number))
                .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
            output_fields = output_fields.checked_add(1).ok_or(Error::InvalidSource)?;
        }
    }

    if output_length > limits.max_output_bytes() {
        return Err(output_limit_error(output_length, limits));
    }
    if output_fields > limits.max_fields() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let root_field_work = view
        .len()
        .checked_mul(2)
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    let rewrite_work = source
        .len()
        .checked_add(output_length)
        .and_then(|work| work.checked_add(root_field_work))
        .and_then(|work| work.checked_add(nested_work))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if rewrite_work > limits.max_rewrite_work() {
        return Err(work_limit_error(rewrite_work, limits));
    }

    let size_rewrite = size_payload
        .map(|payload| rewrite_size_payload(payload, before.size(), after.size(), limits))
        .transpose()?
        .map(|(replacement, _work)| replacement);
    if size_rewrite
        .as_ref()
        .is_some_and(|replacement| replacement.len() != size_payload.map_or(0, <[u8]>::len))
    {
        return Err(Error::Verification);
    }
    let mut output = allocate_bytes(output_length)?;
    for field in view.fields() {
        if field.number() == SHOW_SIZE_FIELD {
            if let Some(replacement) = size_rewrite.as_deref() {
                let header_length = field
                    .raw()
                    .len()
                    .checked_sub(field.payload().len())
                    .ok_or(Error::InvalidSource)?;
                output.extend_from_slice(&field.raw()[..header_length]);
                output.extend_from_slice(replacement);
            } else {
                output.extend_from_slice(field.raw());
            }
        } else if is_scalar_settings_field(field.number())
            && scalar_changed(field.number(), before, after)
        {
            if let Some(value) = scalar_value(field.number(), after) {
                value.append(&mut output, field.number());
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    for field_number in SCALAR_FIELDS {
        let present = seen_scalars & scalar_presence_bit(field_number) != 0;
        if !present
            && scalar_changed(field_number, before, after)
            && let Some(value) = scalar_value(field_number, after)
        {
            value.append(&mut output, field_number);
        }
    }
    if output.len() != output_length {
        return Err(Error::Verification);
    }
    Ok(output)
}

#[allow(
    clippy::float_cmp,
    reason = "semantic equality intentionally treats positive and negative zero as the same setting"
)]
fn inspect_size_payload(source: &[u8], limits: WireLimits) -> Result<usize, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let field_work = view
        .len()
        .checked_mul(2)
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    let rewrite_work = source
        .len()
        .checked_add(source.len())
        .and_then(|work| work.checked_add(field_work))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    let mut width_seen = false;
    let mut height_seen = false;
    for field in view.fields() {
        match field.number() {
            SIZE_WIDTH_FIELD if !width_seen && field.wire_type() == 5 => width_seen = true,
            SIZE_HEIGHT_FIELD if !height_seen && field.wire_type() == 5 => height_seen = true,
            SIZE_WIDTH_FIELD | SIZE_HEIGHT_FIELD => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    if !width_seen || !height_seen {
        return Err(Error::InvalidSource);
    }
    Ok(rewrite_work)
}

#[allow(
    clippy::float_cmp,
    reason = "semantic equality intentionally treats positive and negative zero as the same setting"
)]
fn rewrite_size_payload(
    source: &[u8],
    before: Size,
    after: Size,
    limits: WireLimits,
) -> Result<(Vec<u8>, usize), Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let field_work = view
        .len()
        .checked_mul(2)
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    let rewrite_work = source
        .len()
        .checked_add(source.len())
        .and_then(|work| work.checked_add(field_work))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if rewrite_work > limits.max_rewrite_work() {
        return Err(work_limit_error(rewrite_work, limits));
    }
    let mut output = allocate_bytes(source.len())?;
    let mut width_seen = false;
    let mut height_seen = false;
    for field in view.fields() {
        match field.number() {
            SIZE_WIDTH_FIELD => {
                if width_seen || field.wire_type() != 5 {
                    return Err(Error::InvalidSource);
                }
                width_seen = true;
                if before.width() == after.width() {
                    output.extend_from_slice(field.raw());
                } else {
                    output.extend_from_slice(field.key());
                    output.extend_from_slice(&after.width().to_bits().to_le_bytes());
                }
            },
            SIZE_HEIGHT_FIELD => {
                if height_seen || field.wire_type() != 5 {
                    return Err(Error::InvalidSource);
                }
                height_seen = true;
                if before.height() == after.height() {
                    output.extend_from_slice(field.raw());
                } else {
                    output.extend_from_slice(field.key());
                    output.extend_from_slice(&after.height().to_bits().to_le_bytes());
                }
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if !width_seen || !height_seen || output.len() != source.len() {
        return Err(Error::InvalidSource);
    }
    Ok((output, rewrite_work))
}

const fn is_scalar_settings_field(field_number: u32) -> bool {
    matches!(
        field_number,
        SLIDE_NUMBERS_VISIBLE_FIELD
            | LOOP_PRESENTATION_FIELD
            | MODE_FIELD
            | AUTOPLAY_TRANSITION_DELAY_FIELD
            | AUTOPLAY_BUILD_DELAY_FIELD
            | IDLE_TIMER_ACTIVE_FIELD
            | IDLE_TIMER_DELAY_FIELD
            | AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD
    )
}

const fn scalar_presence_bit(field_number: u32) -> u32 {
    1u32 << field_number
}

const fn scalar_wire_type(field_number: u32) -> u8 {
    match field_number {
        AUTOPLAY_TRANSITION_DELAY_FIELD | AUTOPLAY_BUILD_DELAY_FIELD | IDLE_TIMER_DELAY_FIELD => 1,
        SLIDE_NUMBERS_VISIBLE_FIELD
        | LOOP_PRESENTATION_FIELD
        | MODE_FIELD
        | IDLE_TIMER_ACTIVE_FIELD
        | AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD => 0,
        _ => 7,
    }
}

fn scalar_changed(field_number: u32, before: Settings, after: Settings) -> bool {
    match field_number {
        SLIDE_NUMBERS_VISIBLE_FIELD => {
            before.slide_numbers_visible() != after.slide_numbers_visible()
        },
        LOOP_PRESENTATION_FIELD => before.loop_presentation() != after.loop_presentation(),
        MODE_FIELD => before.mode() != after.mode(),
        AUTOPLAY_TRANSITION_DELAY_FIELD => {
            before.autoplay_transition_delay() != after.autoplay_transition_delay()
        },
        AUTOPLAY_BUILD_DELAY_FIELD => before.autoplay_build_delay() != after.autoplay_build_delay(),
        IDLE_TIMER_ACTIVE_FIELD => before.idle_timer_active() != after.idle_timer_active(),
        IDLE_TIMER_DELAY_FIELD => before.idle_timer_delay() != after.idle_timer_delay(),
        AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD => {
            before.automatically_plays_upon_open() != after.automatically_plays_upon_open()
        },
        _ => false,
    }
}

fn scalar_value(field_number: u32, settings: Settings) -> Option<ScalarValue> {
    match field_number {
        SLIDE_NUMBERS_VISIBLE_FIELD => settings
            .slide_numbers_visible()
            .map(|value| ScalarValue::Varint(u64::from(value))),
        LOOP_PRESENTATION_FIELD => settings
            .loop_presentation()
            .map(|value| ScalarValue::Varint(u64::from(value))),
        MODE_FIELD => settings.mode().map(|mode| {
            let signed = i64::from(mode.as_raw());
            ScalarValue::Varint(u64::from_ne_bytes(signed.to_ne_bytes()))
        }),
        AUTOPLAY_TRANSITION_DELAY_FIELD => seconds_bits(settings.autoplay_transition_delay()),
        AUTOPLAY_BUILD_DELAY_FIELD => seconds_bits(settings.autoplay_build_delay()),
        IDLE_TIMER_ACTIVE_FIELD => settings
            .idle_timer_active()
            .map(|value| ScalarValue::Varint(u64::from(value))),
        IDLE_TIMER_DELAY_FIELD => seconds_bits(settings.idle_timer_delay()),
        AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD => settings
            .automatically_plays_upon_open()
            .map(|value| ScalarValue::Varint(u64::from(value))),
        _ => None,
    }
}

fn seconds_bits(seconds: Option<Seconds>) -> Option<ScalarValue> {
    seconds.map(|value| ScalarValue::Fixed64(value.as_f64().to_bits()))
}

fn allocate_bytes(capacity: usize) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| Error::Allocation { amount: capacity })?;
    Ok(output)
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the internal prepared-source feature adds a semantic-only failure branch"
)]
fn physical_source_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}

fn editable_source_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    let source = physical_source_catalog(package)?;
    if !source.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    Ok(source)
}

fn map_rendering_invalidation_error(
    error: super::rendering_invalidation::RenderingInvalidationError,
) -> Error {
    match error {
        super::rendering_invalidation::RenderingInvalidationError::InvalidSource => {
            Error::InvalidSource
        },
        super::rendering_invalidation::RenderingInvalidationError::Allocation { amount } => {
            Error::Allocation { amount }
        },
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Entries,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::InvalidSource,
    }
}

fn output_limit_error(observed: usize, limits: WireLimits) -> Error {
    Error::LimitExceeded {
        kind: LimitKind::OutputBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_output_bytes()),
    }
}

fn work_limit_error(observed: usize, limits: WireLimits) -> Error {
    Error::LimitExceeded {
        kind: LimitKind::WireWork,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_rewrite_work()),
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
