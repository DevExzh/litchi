//! Immutable transactions for Keynote presentation settings.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::{SourceCatalog, package::EntryEdit};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use thiserror::Error;

use super::{
    Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE, SemanticBudget, SemanticPath,
    decode_show_settings_snapshot, preflight_show, settings_from_show_projection, unique_payload,
};
use crate::{Seconds, Settings, Size};

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
pub enum ShowSettingsLimitKind {
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

impl fmt::Display for ShowSettingsLimitKind {
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
pub enum ShowSettingsError {
    /// The package has no physical source that can back an exact patch.
    #[error("this Keynote source does not support physical show-settings edits")]
    UnsupportedSource,
    /// The requested semantic settings do not satisfy their public invariants.
    #[error("the requested Keynote show settings are invalid")]
    InvalidSettings,
    /// The source package or selected show payload is structurally invalid.
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A finite retained resource ceiling was exceeded.
    #[error("Keynote show-settings {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ShowSettingsLimitKind,
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
pub struct ShowSettingsEdit<'a> {
    source: &'a Package,
    before: Settings,
    settings: Settings,
}

impl<'a> ShowSettingsEdit<'a> {
    fn new(source: &'a Package) -> Result<Self, ShowSettingsError> {
        physical_source_catalog(source)?;
        let before = strict_package_settings(source)?;
        Ok(Self {
            source,
            before,
            settings: before,
        })
    }

    /// Borrow the settings that would be published by this edit.
    #[must_use]
    pub const fn settings(&self) -> &Settings {
        &self.settings
    }

    /// Mutably borrow the settings that would be published by this edit.
    ///
    /// Semantic validation is repeated by [`Self::commit`] before any physical
    /// rewrite work begins.
    #[must_use]
    pub fn settings_mut(&mut self) -> &mut Settings {
        &mut self.settings
    }

    /// Replace the staged settings after validating their semantic invariants.
    ///
    /// # Errors
    ///
    /// Returns [`ShowSettingsError::InvalidSettings`] when `settings` is not a
    /// valid semantic value.
    pub fn set_settings(&mut self, settings: Settings) -> Result<&mut Self, ShowSettingsError> {
        settings
            .validate()
            .map_err(|_error| ShowSettingsError::InvalidSettings)?;
        self.settings = settings;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// An exact semantic no-op reuses the package source allocation. A change
    /// is published only after full source and candidate validation plus a
    /// strict Buffa semantic readback under the retained limits.
    ///
    /// # Errors
    ///
    /// Returns an error without modifying `source` when its physical
    /// provenance, semantic graph, wire payload, resource profile, allocation,
    /// or readback invariant is rejected.
    pub fn commit(self) -> Result<ShowSettingsCommit, ShowSettingsError> {
        self.settings
            .validate()
            .map_err(|_error| ShowSettingsError::InvalidSettings)?;
        let source_catalog = physical_source_catalog(self.source)?;
        let source_bytes = source_catalog.shared_source();
        let source_fingerprint = fingerprint(&source_bytes);

        self.source.validate().map_err(map_read_error)?;
        if strict_package_settings(self.source)? != self.before {
            return Err(ShowSettingsError::InvalidSource);
        }

        if self.before == self.settings {
            return Ok(ShowSettingsCommit {
                package: self.source.snapshot(),
                patch: ShowSettingsPatch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    before: self.before,
                    after: self.settings,
                },
                diagnostics: ShowSettingsDiagnostics::unchanged(),
            });
        }

        if !source_catalog.source_is_exact() {
            return Err(ShowSettingsError::UnsupportedSource);
        }
        let package = rewrite_package_settings(self.source, self.before, self.settings)?;
        let target_bytes = physical_source_catalog(&package)?.shared_source();
        let target_fingerprint = fingerprint(&target_bytes);
        Ok(ShowSettingsCommit {
            package,
            patch: ShowSettingsPatch {
                source_bytes,
                target_bytes,
                source_fingerprint,
                target_fingerprint,
                before: self.before,
                after: self.settings,
            },
            diagnostics: ShowSettingsDiagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible semantic show-settings patch.
///
/// Native identifiers, package member names, and exact source/target bytes are
/// retained privately. Fingerprints are compact diagnostics; exact private
/// byte comparison authorizes patch application.
#[derive(Clone, PartialEq)]
pub struct ShowSettingsPatch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Settings,
    after: Settings,
}

impl fmt::Debug for ShowSettingsPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ShowSettingsPatch")
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl ShowSettingsPatch {
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
        }
    }
}

/// Compact evidence describing one show-settings commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShowSettingsDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl ShowSettingsDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
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

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable show-settings transaction.
#[must_use = "a Keynote show-settings commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ShowSettingsCommit {
    package: Package,
    patch: ShowSettingsPatch,
    diagnostics: ShowSettingsDiagnostics,
}

impl ShowSettingsCommit {
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
    pub const fn patch(&self) -> &ShowSettingsPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ShowSettingsDiagnostics {
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
    /// # Errors
    ///
    /// Returns a typed error when the show payload cannot be projected safely
    /// under this package's retained physical and semantic limits.
    pub fn show_settings(&self) -> Result<Settings, ShowSettingsError> {
        strict_package_settings(self)
    }

    /// Start a focused edit of presentation dimensions and playback settings.
    ///
    /// The transaction's before value comes directly from the strict bounded
    /// Buffa show projection, not from a generated mutable protobuf model.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the package has no physical patch source or
    /// its show payload cannot be projected safely under the retained limits.
    pub fn edit_show_settings(&self) -> Result<ShowSettingsEdit<'_>, ShowSettingsError> {
        ShowSettingsEdit::new(self)
    }

    /// Apply an exact-source-checked show-settings patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original limits before it is published.
    ///
    /// # Errors
    ///
    /// Returns [`ShowSettingsError::PatchConflict`] unless this package is the
    /// exact immutable source captured by `patch`, or another typed validation
    /// error when its retained target cannot be published safely.
    pub fn apply_show_settings(
        &self,
        patch: &ShowSettingsPatch,
    ) -> Result<ShowSettingsCommit, ShowSettingsError> {
        let source_catalog = physical_source_catalog(self)?;
        let source_bytes = source_catalog.shared_source();
        if fingerprint(source_catalog.source_bytes()) != patch.source_fingerprint
            || source_catalog.source_bytes() != patch.source_bytes.as_ref()
            || source_bytes.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(ShowSettingsError::PatchConflict);
        }

        self.validate().map_err(map_read_error)?;
        if strict_package_settings(self)? != patch.before {
            return Err(ShowSettingsError::PatchConflict);
        }

        if patch.is_noop() {
            if patch.source_bytes.as_ref() != patch.target_bytes.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(ShowSettingsError::PatchConflict);
            }
            return Ok(ShowSettingsCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ShowSettingsDiagnostics::unchanged(),
            });
        }

        if !source_catalog.source_is_exact()
            || fingerprint(&patch.target_bytes) != patch.target_fingerprint
        {
            return Err(ShowSettingsError::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target_bytes), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        if strict_package_settings(&candidate)? != patch.after {
            return Err(ShowSettingsError::Verification);
        }
        Ok(ShowSettingsCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ShowSettingsDiagnostics::published(),
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

fn strict_package_settings(package: &Package) -> Result<Settings, ShowSettingsError> {
    let show_identifier = package.root_show_identifier().map_err(map_read_error)?;
    if show_identifier == 0 {
        return Ok(Settings::default());
    }
    let object = package
        .required_object(show_identifier, "Keynote show")
        .map_err(map_read_error)?;
    let payload = unique_payload(&object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")
        .map_err(map_read_error)?;
    strict_payload_settings(package, payload)
}

fn strict_payload_settings(
    package: &Package,
    payload: &[u8],
) -> Result<Settings, ShowSettingsError> {
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
) -> Result<Package, ShowSettingsError> {
    let source_catalog = editable_source_catalog(source)?;
    let show_identifier = source.root_show_identifier().map_err(map_read_error)?;
    if show_identifier == 0 {
        return Err(ShowSettingsError::UnsupportedSource);
    }
    let mut components = source_catalog
        .components()
        .iter()
        .filter(|component| component.archive().object(show_identifier).is_some());
    let component = components.next().ok_or(ShowSettingsError::InvalidSource)?;
    if components.next().is_some() {
        return Err(ShowSettingsError::InvalidSource);
    }
    let component_name = component.name();
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(ShowSettingsError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ShowSettingsError::InvalidSource);
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
    let object = archive
        .object(show_identifier)
        .ok_or(ShowSettingsError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_index, message)| message.type_ == SHOW_MESSAGE_TYPE);
    let (message_index, message) = messages.next().ok_or(ShowSettingsError::InvalidSource)?;
    if messages.next().is_some() {
        return Err(ShowSettingsError::InvalidSource);
    }
    if strict_payload_settings(source, &message.data)? != before {
        return Err(ShowSettingsError::InvalidSource);
    }
    let rewritten = rewrite_show_payload(
        &message.data,
        before,
        after,
        source.wire_limits().map_err(map_wire_error)?,
    )?;
    if strict_payload_settings(source, &rewritten)? != after {
        return Err(ShowSettingsError::Verification);
    }

    archive
        .object_mut(show_identifier)
        .ok_or(ShowSettingsError::InvalidSource)?
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
    let compressed = SnappyStream::compress(&rewritten_archive).map_err(map_core_error)?;
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(component_name, &compressed)],
            physical_limits,
        )
        .map_err(map_archive_error)?;

    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if strict_package_settings(&candidate)? != after {
        return Err(ShowSettingsError::Verification);
    }
    Ok(candidate)
}

fn rewrite_show_payload(
    source: &[u8],
    before: Settings,
    after: Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, ShowSettingsError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output_length = 0usize;
    let mut output_fields = view.len();
    let mut seen_scalars = 0u32;
    let mut saw_size = false;
    let mut size_rewrite = None;
    let mut nested_work = 0usize;

    for field in view.fields() {
        let replacement_length = if field.number() == SHOW_SIZE_FIELD {
            if saw_size || field.wire_type() != 2 {
                return Err(ShowSettingsError::InvalidSource);
            }
            saw_size = true;
            if before.size() == after.size() {
                field.raw().len()
            } else {
                let (replacement, work) =
                    rewrite_size_payload(field.payload(), before.size(), after.size(), limits)?;
                if replacement.len() != field.payload().len() {
                    return Err(ShowSettingsError::Verification);
                }
                nested_work = work;
                let header_length = field
                    .raw()
                    .len()
                    .checked_sub(field.payload().len())
                    .ok_or(ShowSettingsError::InvalidSource)?;
                let length = header_length
                    .checked_add(replacement.len())
                    .ok_or(ShowSettingsError::InvalidSource)?;
                size_rewrite = Some(replacement);
                length
            }
        } else if is_scalar_settings_field(field.number()) {
            let bit = scalar_presence_bit(field.number());
            if seen_scalars & bit != 0 || field.wire_type() != scalar_wire_type(field.number()) {
                return Err(ShowSettingsError::InvalidSource);
            }
            seen_scalars |= bit;
            if scalar_changed(field.number(), before, after) {
                if let Some(value) = scalar_value(field.number(), after) {
                    value.encoded_len(field.number())
                } else {
                    output_fields = output_fields
                        .checked_sub(1)
                        .ok_or(ShowSettingsError::InvalidSource)?;
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
        return Err(ShowSettingsError::InvalidSource);
    }

    for field_number in SCALAR_FIELDS {
        let present = seen_scalars & scalar_presence_bit(field_number) != 0;
        if present != scalar_value(field_number, before).is_some() {
            return Err(ShowSettingsError::InvalidSource);
        }
        if !present
            && scalar_changed(field_number, before, after)
            && let Some(value) = scalar_value(field_number, after)
        {
            output_length = output_length
                .checked_add(value.encoded_len(field_number))
                .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
            output_fields = output_fields
                .checked_add(1)
                .ok_or(ShowSettingsError::InvalidSource)?;
        }
    }

    if output_length > limits.max_output_bytes() {
        return Err(output_limit_error(output_length, limits));
    }
    if output_fields > limits.max_fields() {
        return Err(ShowSettingsError::LimitExceeded {
            kind: ShowSettingsLimitKind::WireFields,
            observed: usize_to_u64(output_fields),
            maximum: usize_to_u64(limits.max_fields()),
        });
    }
    let rewrite_work = view
        .len()
        .checked_add(output_fields)
        .and_then(|work| work.checked_add(nested_work))
        .ok_or_else(|| work_limit_error(usize::MAX, limits))?;
    if rewrite_work > limits.max_rewrite_work() {
        return Err(work_limit_error(rewrite_work, limits));
    }

    let mut output = allocate_bytes(output_length)?;
    for field in view.fields() {
        if field.number() == SHOW_SIZE_FIELD {
            if let Some(replacement) = size_rewrite.as_deref() {
                let header_length = field
                    .raw()
                    .len()
                    .checked_sub(field.payload().len())
                    .ok_or(ShowSettingsError::InvalidSource)?;
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
        return Err(ShowSettingsError::Verification);
    }
    Ok(output)
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
) -> Result<(Vec<u8>, usize), ShowSettingsError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let rewrite_work = view
        .len()
        .checked_mul(2)
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
                    return Err(ShowSettingsError::InvalidSource);
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
                    return Err(ShowSettingsError::InvalidSource);
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
        return Err(ShowSettingsError::InvalidSource);
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

fn allocate_bytes(capacity: usize) -> Result<Vec<u8>, ShowSettingsError> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_allocation| ShowSettingsError::Allocation { amount: capacity })?;
    Ok(output)
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the internal prepared-source feature adds a semantic-only failure branch"
)]
fn physical_source_catalog(package: &Package) -> Result<&SourceCatalog, ShowSettingsError> {
    #[cfg(feature = "internal-iwork-source")]
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ShowSettingsError::UnsupportedSource),
    }
    #[cfg(not(feature = "internal-iwork-source"))]
    {
        let PhysicalSource::Package(source) = &package.state.source;
        Ok(source)
    }
}

fn editable_source_catalog(package: &Package) -> Result<&SourceCatalog, ShowSettingsError> {
    let source = physical_source_catalog(package)?;
    if !source.source_is_exact() {
        return Err(ShowSettingsError::UnsupportedSource);
    }
    Ok(source)
}

fn map_read_error(error: ReadError) -> ShowSettingsError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ShowSettingsError::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => ShowSettingsLimitKind::Entries,
                super::SemanticLimitKind::Slides => ShowSettingsLimitKind::Slides,
                super::SemanticLimitKind::References => ShowSettingsLimitKind::References,
                super::SemanticLimitKind::TextStorages => ShowSettingsLimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => ShowSettingsLimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => ShowSettingsLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ShowSettingsError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ShowSettingsLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ShowSettingsLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ShowSettingsLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ShowSettingsLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ShowSettingsError::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => ShowSettingsError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ShowSettingsError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ShowSettingsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ShowSettingsLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => ShowSettingsLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => ShowSettingsLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    ShowSettingsLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => ShowSettingsLimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ShowSettingsError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => ShowSettingsError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_core_error(error: litchi_iwa_core::Error) -> ShowSettingsError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ShowSettingsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ShowSettingsLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes => ShowSettingsLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => ShowSettingsLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ShowSettingsLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => ShowSettingsLimitKind::EntryBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ShowSettingsError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => ShowSettingsError::InvalidSource,
    }
}

#[allow(
    clippy::needless_pass_by_value,
    reason = "used directly as a Result::map_err conversion"
)]
fn map_wire_error(error: litchi_iwa_common::Error) -> ShowSettingsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ShowSettingsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ShowSettingsLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ShowSettingsLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ShowSettingsLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => ShowSettingsLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ShowSettingsLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ShowSettingsError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ShowSettingsError::InvalidSource,
    }
}

fn output_limit_error(observed: usize, limits: WireLimits) -> ShowSettingsError {
    ShowSettingsError::LimitExceeded {
        kind: ShowSettingsLimitKind::OutputBytes,
        observed: usize_to_u64(observed),
        maximum: usize_to_u64(limits.max_output_bytes()),
    }
}

fn work_limit_error(observed: usize, limits: WireLimits) -> ShowSettingsError {
    ShowSettingsError::LimitExceeded {
        kind: ShowSettingsLimitKind::WireWork,
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
