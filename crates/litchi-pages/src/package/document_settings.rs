//! Immutable, exact-source transactions for combined Pages document settings.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireView};
use litchi_iwa_core::RawMessage;
use litchi_iwa_protos::pages_document_settings_codec::{
    self, DecodeOptions as SettingsDecodeOptions, DocumentSettingsSnapshot, WireResourceLimit,
};
use thiserror::Error as ThisError;

use super::{Package, PackageError, page_layout};
use crate::{document_settings::Settings, footnote};

const DOCUMENT_IDENTIFIER: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const SETTINGS_MESSAGE_TYPE: u32 = 10_012;
const SETTINGS_REFERENCE_FIELD: u32 = 7;
const SETTINGS_FIELDS: [u32; 10] = [1, 2, 3, 9, 10, 30, 31, 32, 33, 34];

/// A finite resource enforced while reading or publishing document settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete edited package output bytes.
    OutputBytes,
    /// Package members retained by the transaction.
    Entries,
    /// Bytes retained by one package member.
    EntryBytes,
    /// Aggregate bytes retained by package members.
    TotalEntryBytes,
    /// Package names and structural metadata.
    PackageBytes,
    /// Bytes in one decoded component payload.
    PayloadBytes,
    /// Aggregate decoded component payload bytes.
    TotalPayloadBytes,
    /// Component objects inspected by the transaction.
    PayloadObjects,
    /// Component messages inspected by the transaction.
    PayloadMessages,
    /// Component framing or metadata items inspected by the transaction.
    PayloadItems,
    /// Bytes inspected by the settings projection.
    WireBytes,
    /// Fields inspected by the settings projection.
    WireFields,
    /// Nesting inspected by the settings projection.
    WireNesting,
    /// Aggregate settings rewrite work.
    WireWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "package entries",
            Self::EntryBytes => "package entry bytes",
            Self::TotalEntryBytes => "total package entry bytes",
            Self::PackageBytes => "package metadata bytes",
            Self::PayloadBytes => "component payload bytes",
            Self::TotalPayloadBytes => "total component payload bytes",
            Self::PayloadObjects => "component objects",
            Self::PayloadMessages => "component messages",
            Self::PayloadItems => "component metadata items",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// Failure from a semantic document-settings read or immutable transaction.
///
/// Display output never exposes package bytes, member names, object
/// identifiers, field values, or retained patch artifacts.
#[derive(Debug, ThisError, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The requested semantic settings violate an archive-free invariant.
    #[error("invalid Pages document settings: {0}")]
    InvalidSettings(#[source] footnote::Error),
    /// The source cannot publish a preservation-safe changed edit.
    #[error("the Pages package source does not support exact document-settings editing")]
    UnsupportedSource,
    /// The rooted native settings or derived cache cannot be handled safely.
    #[error("the Pages source has no unambiguous editable document settings")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Pages document-settings {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: LimitKind,
        /// Observed or requested resource amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Pages document-settings transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
    },
    /// Complete candidate reopening did not reproduce the requested state.
    #[error("the edited Pages document settings failed semantic verification")]
    Verification,
    /// The supplied patch was not created from this exact package artifact.
    #[error("the Pages document-settings patch does not match the exact source package")]
    PatchConflict,
}

/// Combined document settings staged against one immutable package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    location: SettingsLocation,
    before: Settings,
    settings: Settings,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("settings", &self.settings)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the lossless combined settings that would be published.
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.settings
    }

    /// Replace the staged valid settings, returning the edit for chaining.
    ///
    /// # Costs
    ///
    /// Assignment is constant-time and does not inspect the package.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.settings = settings;
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// Exact semantic no-ops reuse the source allocation before inspecting
    /// cache ownership. Changed edits require an exact flat source, invalidate
    /// the rooted layout cache, delete root previews, and fully reopen the
    /// complete candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed error for invalid settings, stale or ambiguous rooted
    /// state, unsupported source shape, finite resource exhaustion, allocation
    /// failure, or failed complete readback.
    ///
    /// # Costs
    ///
    /// A no-op validates only the staged value and reuses the exact source
    /// allocation. A change scans the rooted cache graph, rewrites at most two
    /// component payloads, reassembles the package, and fully reopens and
    /// verifies the candidate under the retained source limits.
    pub fn commit(self) -> Result<Commit, Error> {
        self.settings.validate().map_err(Error::InvalidSettings)?;
        let catalog = &self.source.state.source;
        let source = catalog.shared_source();
        let source_fingerprint = page_layout::fingerprint(&source);
        if self.before == self.settings {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    source: Arc::clone(&source),
                    target: source,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    before: self.before,
                    after: self.settings,
                    source_layout_state: None,
                    target_layout_state: None,
                    source_preview_count: 0,
                    target_preview_count: 0,
                    touched_components: 0,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let view_location =
            page_layout::view_state_location(self.source).map_err(map_page_layout_error)?;
        let source_layout_state = view_location
            .as_ref()
            .map(|location| location.layout_identifier);
        let source_preview_count = page_layout::preview_count(self.source);
        let (package, touched_components) = rewrite_document_settings(
            self.source,
            &self.location,
            self.before,
            self.settings,
            view_location.as_ref(),
        )?;
        let target = package.state.source.shared_source();
        let target_fingerprint = page_layout::fingerprint(&target);
        let target_layout_state = None;
        let target_preview_count = 0;
        let deleted_previews = source_preview_count.saturating_sub(target_preview_count);
        Ok(Commit {
            package,
            patch: Patch {
                source,
                target,
                source_fingerprint,
                target_fingerprint,
                before: self.before,
                after: self.settings,
                source_layout_state,
                target_layout_state,
                source_preview_count,
                target_preview_count,
                touched_components,
            },
            diagnostics: Diagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// A reversible patch bound to exact source and target package artifacts.
///
/// Exact artifacts and ownership facts remain private. Fingerprints are
/// diagnostic only and never replace exact artifact matching.
#[derive(Clone, PartialEq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Settings,
    after: Settings,
    source_layout_state: Option<u64>,
    target_layout_state: Option<u64>,
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

    /// Return the semantic settings represented by the patch target.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether this patch retains both semantic settings and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source, &self.target)
                || self.source.as_ref() == self.target.as_ref())
    }

    /// Return the exact patch from the target artifact back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            before: self.after,
            after: self.before,
            source_layout_state: self.target_layout_state,
            target_layout_state: self.source_layout_state,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            touched_components: self.touched_components,
        }
    }
}

/// Compact evidence describing one document-settings commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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

    /// Return whether the package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten component payloads.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of root preview entries deleted by the edit.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete target package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable document-settings transaction.
#[must_use = "a document-settings commit contains the validated package snapshot"]
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
    /// Read the lossless combined document and footnote formatter settings.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the rooted settings graph is ambiguous,
    /// malformed, semantically invalid, or exceeds a finite decode limit.
    ///
    /// # Costs
    ///
    /// Scans the component index, then strictly decodes the bounded document
    /// root and selected settings payload without modifying the package.
    pub fn document_settings(&self) -> Result<Settings, Error> {
        document_settings(self)
    }

    /// Start a document-wide immutable settings edit.
    ///
    /// # Errors
    ///
    /// Returns the same typed read errors as [`Package::document_settings`].
    ///
    /// # Costs
    ///
    /// Performs one settings read and retains copyable semantic values, a
    /// bounded private rooted location, and a borrow of this immutable package.
    pub fn edit_document_settings(&self) -> Result<Edit<'_>, Error> {
        let location = settings_location(self)?;
        let before = location.settings;
        Ok(Edit {
            source: self,
            location,
            before,
            settings: before,
        })
    }

    /// Apply an exact-source-checked document-settings patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] unless this package exactly matches
    /// the patch source, or another typed error if the retained target cannot
    /// be safely reopened and completely verified.
    ///
    /// # Costs
    ///
    /// Exact matching may scan the complete package bytes. A no-op reuses the
    /// source allocation; a changed apply reopens the complete retained target
    /// and verifies settings, cache invalidation, previews, and document
    /// semantics under the original source limits.
    pub fn apply_document_settings(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = &self.state.source;
        let source = catalog.shared_source();
        let exact_source = Arc::ptr_eq(&source, &patch.source)
            || (page_layout::fingerprint(catalog.source_bytes()) == patch.source_fingerprint
                && catalog.source_bytes() == patch.source.as_ref());
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
        if document_settings(self)? != patch.before
            || page_layout::view_state_layout_identifier(self).map_err(map_page_layout_error)?
                != patch.source_layout_state
            || page_layout::preview_count(self) != patch.source_preview_count
            || !catalog.source_is_exact()
            || page_layout::fingerprint(&patch.target) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target),
            catalog.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        verify_candidate(
            self,
            &candidate,
            patch.after,
            patch.target_layout_state,
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

#[derive(Debug, Clone)]
struct SettingsLocation {
    component_name: String,
    object_identifier: u64,
    settings: Settings,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scalar {
    Bool(bool),
    Int32(i32),
}

fn document_settings(package: &Package) -> Result<Settings, Error> {
    settings_facts(package).map(|(_component, _identifier, settings)| settings)
}

fn settings_location(package: &Package) -> Result<SettingsLocation, Error> {
    let (component_name, object_identifier, settings) = settings_facts(package)?;
    Ok(SettingsLocation {
        component_name: try_owned(component_name)?,
        object_identifier,
        settings,
    })
}

fn settings_facts(package: &Package) -> Result<(&str, u64, Settings), Error> {
    let limits = wire_limits(package)?;
    let (_root_component, root) = page_layout::object_location(package, DOCUMENT_IDENTIFIER)
        .map_err(map_page_layout_error)?
        .ok_or(Error::InvalidSource)?;
    let (root_message_index, root_message) =
        page_layout::unique_message(root, DOCUMENT_MESSAGE_TYPE).map_err(map_page_layout_error)?;
    page_layout::validate_selected_metadata(root, root_message_index)
        .map_err(map_page_layout_error)?;
    let settings_identifier = strict_settings_reference(&root_message.data, limits)?;
    page_layout::validate_reference_metadata(
        root,
        root_message_index,
        settings_identifier,
        &[SETTINGS_REFERENCE_FIELD],
    )
    .map_err(map_page_layout_error)?;
    let (component_name, object) = page_layout::object_location(package, settings_identifier)
        .map_err(map_page_layout_error)?
        .ok_or(Error::InvalidSource)?;
    let (message_index, message) = page_layout::unique_message(object, SETTINGS_MESSAGE_TYPE)
        .map_err(map_page_layout_error)?;
    page_layout::validate_selected_metadata(object, message_index)
        .map_err(map_page_layout_error)?;
    let settings = strict_settings(&root_message.data, &message.data, limits)?;
    Ok((component_name, settings_identifier, settings))
}

fn strict_settings_reference(payload: &[u8], limits: WireLimits) -> Result<u64, Error> {
    pages_document_settings_codec::decode_document_settings_reference(
        payload,
        decode_options(limits)?,
    )
    .map(|reference| reference.identifier().get())
    .map_err(map_codec_error)
}

fn strict_settings(
    root_payload: &[u8],
    settings_payload: &[u8],
    limits: WireLimits,
) -> Result<Settings, Error> {
    let snapshot = pages_document_settings_codec::decode_document_settings(
        root_payload,
        settings_payload,
        decode_options(limits)?,
    )
    .map_err(map_codec_error)?;
    projected_settings(snapshot)
}

fn projected_settings(snapshot: DocumentSettingsSnapshot) -> Result<Settings, Error> {
    let values = snapshot.document_options();
    let options = crate::document_options::Options::new(
        values.body(),
        values.headers(),
        values.footers(),
        values.facing_pages(),
        values.hyphenation(),
        values.use_ligatures(),
    );
    let footnotes = footnote::Settings {
        kind: snapshot.footnote_kind().map(footnote::Kind::from_raw),
        format: snapshot.footnote_format().map(footnote::Format::from_raw),
        numbering: snapshot
            .footnote_numbering()
            .map(footnote::Numbering::from_raw),
        gap: snapshot
            .footnote_gap()
            .map(|value| u32::try_from(value).map_err(|_error| Error::InvalidSource))
            .transpose()?
            .map(footnote::Gap::new)
            .transpose()
            .map_err(Error::InvalidSettings)?,
    };
    Settings::new(options, footnotes).map_err(Error::InvalidSettings)
}

fn decode_options(limits: WireLimits) -> Result<SettingsDecodeOptions, Error> {
    let recursion = u32::try_from(limits.max_nesting()).map_err(|_error| Error::LimitExceeded {
        kind: LimitKind::WireNesting,
        observed: usize_to_u64(limits.max_nesting()),
        maximum: u64::from(u32::MAX),
    })?;
    Ok(SettingsDecodeOptions::new(
        limits.max_input_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn rewrite_document_settings(
    source: &Package,
    location: &SettingsLocation,
    before: Settings,
    after: Settings,
    view_location: Option<&page_layout::ViewStateLocation>,
) -> Result<(Package, usize), Error> {
    let catalog = &source.state.source;
    let shared_component =
        view_location.is_some_and(|view| view.component_name == location.component_name);
    let settings_compressed = rewrite_settings_component(
        source,
        location,
        before,
        after,
        shared_component.then_some(view_location).flatten(),
    )?;
    let view_compressed = if shared_component {
        None
    } else {
        view_location
            .map(|view| rewrite_view_component(source, view))
            .transpose()?
    };

    let mut compressed = Vec::new();
    compressed
        .try_reserve_exact(usize::from(view_compressed.is_some()) + 1)
        .map_err(|_allocation| Error::Allocation { amount: 2 })?;
    compressed.push((location.component_name.as_str(), settings_compressed));
    if let (Some(view), Some(bytes)) = (view_location, view_compressed) {
        compressed.push((view.component_name.as_str(), bytes));
    }
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(compressed.len())
        .map_err(|_allocation| Error::Allocation {
            amount: compressed.len(),
        })?;
    for (name, bytes) in &compressed {
        edits.push(EntryEdit::new(name, bytes));
    }
    let mut deleted_previews = Vec::new();
    deleted_previews
        .try_reserve_exact(page_layout::PREVIEW_ENTRY_NAMES.len())
        .map_err(|_allocation| Error::Allocation {
            amount: page_layout::PREVIEW_ENTRY_NAMES.len(),
        })?;
    for name in page_layout::PREVIEW_ENTRY_NAMES {
        if catalog.package().iter().any(|entry| entry.name() == name) {
            deleted_previews.push(name);
        }
    }
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deleted_previews, catalog.limits())
        .map_err(map_archive_error)?;
    let touched_components = compressed.len();
    drop(edits);
    drop(deleted_previews);
    drop(compressed);
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), catalog.limits())
            .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    verify_candidate(source, &candidate, after, None, 0)?;
    Ok((candidate, touched_components))
}

fn rewrite_settings_component(
    source: &Package,
    location: &SettingsLocation,
    before: Settings,
    after: Settings,
    view_location: Option<&page_layout::ViewStateLocation>,
) -> Result<Vec<u8>, Error> {
    let (mut archive, limits) = page_layout::editable_archive(source, &location.component_name)
        .map_err(map_page_layout_error)?;
    let root_payload = root_payload(source)?;
    {
        let object = archive
            .object_mut(location.object_identifier)
            .ok_or(Error::InvalidSource)?;
        let (message_index, message) = page_layout::unique_message(object, SETTINGS_MESSAGE_TYPE)
            .map_err(map_page_layout_error)?;
        page_layout::validate_selected_metadata(object, message_index)
            .map_err(map_page_layout_error)?;
        if strict_settings(root_payload, &message.data, wire_limits(source)?)? != before {
            return Err(Error::InvalidSource);
        }
        let rewritten =
            rewrite_settings_payload(&message.data, before, after, wire_limits(source)?)?;
        if strict_settings(root_payload, &rewritten, wire_limits(source)?)? != after {
            return Err(Error::Verification);
        }
        object
            .replace_message_preserving_header_with_limits(
                message_index,
                RawMessage {
                    type_: SETTINGS_MESSAGE_TYPE,
                    data: rewritten,
                },
                limits,
            )
            .map_err(map_core_error)?;
    }
    if let Some(view) = view_location {
        page_layout::invalidate_view_state_in_archive(source, &mut archive, view, limits)
            .map_err(map_page_layout_error)?;
    }
    page_layout::compress_archive(archive, limits).map_err(map_page_layout_error)
}

fn rewrite_view_component(
    source: &Package,
    view: &page_layout::ViewStateLocation,
) -> Result<Vec<u8>, Error> {
    let (mut archive, limits) = page_layout::editable_archive(source, &view.component_name)
        .map_err(map_page_layout_error)?;
    page_layout::invalidate_view_state_in_archive(source, &mut archive, view, limits)
        .map_err(map_page_layout_error)?;
    page_layout::compress_archive(archive, limits).map_err(map_page_layout_error)
}

fn root_payload(package: &Package) -> Result<&[u8], Error> {
    let (_component, object) = page_layout::object_location(package, DOCUMENT_IDENTIFIER)
        .map_err(map_page_layout_error)?
        .ok_or(Error::InvalidSource)?;
    let (_index, message) = page_layout::unique_message(object, DOCUMENT_MESSAGE_TYPE)
        .map_err(map_page_layout_error)?;
    Ok(&message.data)
}

fn rewrite_settings_payload(
    source: &[u8],
    before: Settings,
    after: Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut selected_lengths = [None; 10];
    for field in view.fields() {
        let Some(index) = SETTINGS_FIELDS
            .iter()
            .position(|number| *number == field.number())
        else {
            continue;
        };
        if selected_lengths[index].replace(field.raw().len()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let mut output_length = source.len();
    let mut output_fields = view.len();
    for (index, field_number) in SETTINGS_FIELDS.into_iter().enumerate() {
        let before_value = settings_scalar(before, field_number);
        let after_value = settings_scalar(after, field_number);
        if before_value == after_value {
            continue;
        }
        if let Some(length) = selected_lengths[index] {
            output_length = output_length
                .checked_sub(length)
                .ok_or(Error::InvalidSource)?;
            output_fields = output_fields.checked_sub(1).ok_or(Error::InvalidSource)?;
        }
        if let Some(value) = after_value {
            output_length = output_length
                .checked_add(encoded_scalar_length(field_number, value))
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
    let work = source
        .len()
        .checked_add(output_length)
        .and_then(|amount| amount.checked_add(view.len()))
        .and_then(|amount| amount.checked_add(view.len()))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if work > limits.max_rewrite_work() {
        return Err(work_limit_error(work, limits));
    }

    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_allocation| Error::Allocation {
            amount: output_length,
        })?;
    let mut emitted = [false; 10];
    for field in view.fields() {
        let Some(index) = SETTINGS_FIELDS
            .iter()
            .position(|number| *number == field.number())
        else {
            output.extend_from_slice(field.raw());
            continue;
        };
        emitted[index] = true;
        let before_value = settings_scalar(before, field.number());
        let after_value = settings_scalar(after, field.number());
        if before_value == after_value {
            output.extend_from_slice(field.raw());
        } else if let Some(value) = after_value {
            append_scalar(&mut output, field.number(), value);
        }
    }
    for (index, field_number) in SETTINGS_FIELDS.into_iter().enumerate() {
        if !emitted[index]
            && let Some(value) = settings_scalar(after, field_number)
        {
            append_scalar(&mut output, field_number, value);
        }
    }
    if output.len() != output_length {
        return Err(Error::Verification);
    }
    Ok(output)
}

fn settings_scalar(settings: Settings, field_number: u32) -> Option<Scalar> {
    let options = settings.options();
    let footnotes = settings.footnotes();
    match field_number {
        1 => options.body_enabled().map(Scalar::Bool),
        2 => options.headers_enabled().map(Scalar::Bool),
        3 => options.footers_enabled().map(Scalar::Bool),
        9 => options.automatic_hyphenation().map(Scalar::Bool),
        10 => options.ligatures_enabled().map(Scalar::Bool),
        30 => footnotes
            .kind
            .map(footnote::Kind::as_raw)
            .map(Scalar::Int32),
        31 => footnotes
            .format
            .map(footnote::Format::as_raw)
            .map(Scalar::Int32),
        32 => footnotes
            .numbering
            .map(footnote::Numbering::as_raw)
            .map(Scalar::Int32),
        33 => footnotes
            .gap
            .and_then(|gap| i32::try_from(gap.points()).ok())
            .map(Scalar::Int32),
        34 => options.facing_pages().map(Scalar::Bool),
        _ => None,
    }
}

fn encoded_scalar_length(field_number: u32, scalar: Scalar) -> usize {
    encoded_len(u64::from(field_number) << 3).saturating_add(encoded_len(scalar_varint(scalar)))
}

fn append_scalar(output: &mut Vec<u8>, field_number: u32, scalar: Scalar) {
    encode_varint_into(output, u64::from(field_number) << 3);
    encode_varint_into(output, scalar_varint(scalar));
}

fn scalar_varint(scalar: Scalar) -> u64 {
    match scalar {
        Scalar::Bool(value) => u64::from(value),
        Scalar::Int32(value) => i64::from(value).cast_unsigned(),
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

fn verify_candidate(
    source: &Package,
    candidate: &Package,
    expected: Settings,
    expected_layout_state: Option<u64>,
    expected_preview_count: usize,
) -> Result<(), Error> {
    if document_settings(candidate)? != expected
        || page_layout::view_state_layout_identifier(candidate).map_err(map_page_layout_error)?
            != expected_layout_state
        || page_layout::preview_count(candidate) != expected_preview_count
        || source.stats() != candidate.stats()
        || source.sections().len() != candidate.sections().len()
    {
        return Err(Error::Verification);
    }
    for (before, after) in source.sections().iter().zip(candidate.sections()) {
        if before.name() != after.name()
            || before.section_type() != after.section_type()
            || before.heading() != after.heading()
            || before.paragraphs() != after.paragraphs()
            || before.text_storages() != after.text_storages()
            || before.page_count() != after.page_count()
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn wire_limits(package: &Package) -> Result<WireLimits, Error> {
    page_layout::wire_limits(package).map_err(map_page_layout_error)
}

fn try_owned(source: &str) -> Result<String, Error> {
    page_layout::try_owned(source).map_err(map_page_layout_error)
}

fn map_page_layout_error(error: page_layout::PageLayoutError) -> Error {
    match error {
        page_layout::PageLayoutError::UnsupportedSource => Error::UnsupportedSource,
        page_layout::PageLayoutError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: map_page_layout_limit(kind),
            observed,
            maximum,
        },
        page_layout::PageLayoutError::Allocation { amount } => Error::Allocation { amount },
        page_layout::PageLayoutError::Verification => Error::Verification,
        page_layout::PageLayoutError::InvalidLayout(_)
        | page_layout::PageLayoutError::InvalidSource
        | page_layout::PageLayoutError::PatchConflict => Error::InvalidSource,
    }
}

fn map_page_layout_limit(kind: page_layout::PageLayoutLimitKind) -> LimitKind {
    match kind {
        page_layout::PageLayoutLimitKind::InputBytes => LimitKind::InputBytes,
        page_layout::PageLayoutLimitKind::OutputBytes => LimitKind::OutputBytes,
        page_layout::PageLayoutLimitKind::Entries => LimitKind::Entries,
        page_layout::PageLayoutLimitKind::EntryBytes => LimitKind::EntryBytes,
        page_layout::PageLayoutLimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
        page_layout::PageLayoutLimitKind::PackageBytes => LimitKind::PackageBytes,
        page_layout::PageLayoutLimitKind::PayloadBytes => LimitKind::PayloadBytes,
        page_layout::PageLayoutLimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        page_layout::PageLayoutLimitKind::PayloadObjects => LimitKind::PayloadObjects,
        page_layout::PageLayoutLimitKind::PayloadMessages => LimitKind::PayloadMessages,
        page_layout::PageLayoutLimitKind::PayloadItems => LimitKind::PayloadItems,
        page_layout::PageLayoutLimitKind::WireBytes => LimitKind::WireBytes,
        page_layout::PageLayoutLimitKind::WireFields => LimitKind::WireFields,
        page_layout::PageLayoutLimitKind::WireNesting => LimitKind::WireNesting,
        page_layout::PageLayoutLimitKind::WireWork => LimitKind::WireWork,
    }
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Archive(archive_error) => map_archive_error(archive_error),
        PackageError::SectionNamesTooLarge { observed, limit } => Error::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::NotPages => Error::UnsupportedSource,
        PackageError::PayloadLimit { observed, limit } => Error::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::ObjectLimit { observed, limit } => Error::LimitExceeded {
            kind: LimitKind::PayloadObjects,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::Io(_)
        | PackageError::Detection(_)
        | PackageError::InvalidFormat(_)
        | PackageError::Semantic(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "Result::map_err supplies the owned codec error directly."
)]
fn map_codec_error(error: pages_document_settings_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    match error.wire_resource_limit() {
        Some(WireResourceLimit::Bytes { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => Error::LimitExceeded {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        _ => Error::InvalidSource,
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
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::PackageBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Reassembly(_) => Error::UnsupportedSource,
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::InvalidBundle(_) => Error::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "Result::map_err supplies the owned component error directly."
)]
fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::PayloadItems,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    LimitKind::PayloadBytes
                },
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
    reason = "Result::map_err supplies the owned wire error directly."
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
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
