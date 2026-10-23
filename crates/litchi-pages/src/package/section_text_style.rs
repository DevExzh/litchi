//! Character styling of body text spans (`IWORK_API_REQUESTS.md` request 13,
//! record 0309).
//!
//! A Pages body is one `TSWP.StorageArchive` whose `table_char_style`
//! (field 8) names the character style of each run. Styling a span of a
//! section gives every run inside the span a new
//! `TSWP.CharacterStyleArchive` variation of the run's style (or a
//! parentless variation) that overrides the requested properties, so every
//! other property stays inherited. The variations are inserted beside the
//! body storage's stylesheet, appended to its styles (and to the parent's
//! children), and registered in `Index/Metadata.iwa` (object UUIDs, the
//! body component's references to them, the issued-identifier watermark).
//! Only the body storage, the stylesheet and the registry are rewritten;
//! every other member keeps its exact bytes. The patch retains both exact
//! artifacts, so its inverse restores the source byte for byte.

use std::collections::BTreeSet;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    Batch, ComponentDescriptor, ComponentSelector, ExternalReferenceAddition, ObjectUuidAddition,
    ObjectUuidDescriptor, PackageMetadataVisitor, RewriteError, RewriteLimit, RewriteOptions,
    UuidBits, inspect_package_metadata_with_visitor, rewrite_package_metadata,
};
use litchi_iwa_text::TextSpan;
use litchi_iwa_text_wire::CharacterProperties;
use thiserror::Error as ThisError;

use super::section_text::{native_body, native_section_range};
use super::{Package, find_object};
use crate::SectionSelector;

/// `TSWP.StorageArchive`.
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
/// `TSWP.CharacterStyleArchive`.
const CHARACTER_STYLE_MESSAGE_TYPE: u32 = 2_021;
/// `TSS.StylesheetArchive`.
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
/// `TSP.PackageMetadata`.
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_ENTRY_NAME: &str = "Index/Metadata.iwa";
/// The message versions of a new character style when the stylesheet has
/// none to copy.
const CHARACTER_STYLE_VERSION: [u32; 3] = [1, 0, 5];

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Native object identifiers.
    Identifiers,
    /// Bytes of the package registry.
    RegistryBytes,
    /// Wire fields inspected or emitted in the package registry.
    RegistryFields,
    /// Work performed while inspecting or rewriting the package registry.
    RegistryWork,
    /// Nested wire messages in the package registry.
    RegistryNesting,
    /// Components inspected in the package registry.
    RegistryComponents,
    /// References inspected in the package registry.
    RegistryReferences,
    /// Additions staged in the package registry.
    RegistryAdditions,
    /// A bounded character-style wire operation.
    Wire(litchi_iwa_common::LimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Identifiers => "object identifiers",
            Self::RegistryBytes => "registry bytes",
            Self::RegistryFields => "registry fields",
            Self::RegistryWork => "registry work",
            Self::RegistryNesting => "registry nesting",
            Self::RegistryComponents => "registry components",
            Self::RegistryReferences => "registry references",
            Self::RegistryAdditions => "registry additions",
            Self::Wire(kind) => return kind.fmt(formatter),
        })
    }
}

/// A typed, content-redacted text-style failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted text-style source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} text-style items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("text-style candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("text-style patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The section selector resolved no section.
    #[error("section selector did not resolve")]
    SectionNotFound,
    /// The span is empty or past the section's text.
    #[error("text span outside the section")]
    InvalidSpan,
    /// The style sets no property.
    #[error("the style sets no property")]
    EmptyStyle,
    /// Commit was requested before staging a span.
    #[error("no text style is staged")]
    NoStagedOperation,
    /// The body storage has no stylesheet, or its styles live in another
    /// component than the stylesheet.
    #[error("the body's styles are outside what this transaction rewrites")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support text styling")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SectionNotFound { .. } => RefusalKind::NotFound,
            Self::InvalidSpan { .. } | Self::EmptyStyle { .. } => RefusalKind::InvalidValue,
            Self::NoStagedOperation { .. } => RefusalKind::UnsupportedOperation,
            Self::UnsupportedDependency { .. } | Self::UnsupportedSource { .. } => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<litchi_iwa_core::Error> for Error {
    fn from(error: litchi_iwa_core::Error) -> Self {
        match error {
            litchi_iwa_core::Error::Allocation { requested, .. } => {
                Self::Allocation { amount: requested }
            },
            _ => Self::InvalidSource,
        }
    }
}

impl From<litchi_iwa_text_wire::CharStyleError> for Error {
    fn from(error: litchi_iwa_text_wire::CharStyleError) -> Self {
        match error {
            litchi_iwa_text_wire::CharStyleError::Span => Self::Refused(Refusal::InvalidSpan),
            litchi_iwa_text_wire::CharStyleError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: LimitKind::Wire(kind),
                observed: observed as u64,
                maximum: maximum as u64,
            },
            litchi_iwa_text_wire::CharStyleError::Allocation { amount } => {
                Self::Allocation { amount }
            },
            _ => Self::InvalidSource,
        }
    }
}

fn map_archive(error: &litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            Error::Allocation { amount: *amount }
        },
        _ => Error::Refused(Refusal::UnsupportedSource),
    }
}

fn map_rewrite(error: &RewriteError) -> Error {
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    let (kind, observed, maximum) = match error.resource_limit() {
        Some(
            RewriteLimit::InputBytes { observed, maximum }
            | RewriteLimit::OutputBytes { observed, maximum },
        ) => (LimitKind::RegistryBytes, observed as u64, maximum as u64),
        Some(RewriteLimit::Fields { observed, maximum }) => {
            (LimitKind::RegistryFields, observed as u64, maximum as u64)
        },
        Some(RewriteLimit::Work { observed, maximum }) => {
            (LimitKind::RegistryWork, observed as u64, maximum as u64)
        },
        Some(RewriteLimit::Nesting { observed, maximum }) => (
            LimitKind::RegistryNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        Some(RewriteLimit::Components { observed, maximum }) => (
            LimitKind::RegistryComponents,
            observed as u64,
            maximum as u64,
        ),
        Some(RewriteLimit::References { observed, maximum }) => (
            LimitKind::RegistryReferences,
            observed as u64,
            maximum as u64,
        ),
        Some(RewriteLimit::Additions { observed, maximum }) => (
            LimitKind::RegistryAdditions,
            observed as u64,
            maximum as u64,
        ),
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed,
        maximum,
    }
}

/// One run of a section's text and the character properties its style
/// sets itself.
#[derive(Debug, Clone, PartialEq)]
pub struct Run {
    span: TextSpan,
    style: CharacterProperties,
}

impl Run {
    /// The run's UTF-16 span within the section.
    #[must_use]
    pub const fn span(&self) -> TextSpan {
        self.span
    }

    /// The properties the run's character style and its parent chain set
    /// (empty for unstyled text).
    #[must_use]
    pub const fn style(&self) -> &CharacterProperties {
        &self.style
    }
}

/// One immutable edit that styles spans of one section's text, in order.
pub struct Edit<'a> {
    source: &'a Package,
    section: Position,
    spans: Vec<(TextSpan, CharacterProperties)>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("section", &self.section)
            .field("spans", &self.spans.len())
            .finish()
    }
}

/// A reversible, exact-source text-style patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact artifact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one text-style publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    created_styles: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.created_styles != 0
    }

    /// Number of character-style variations created.
    #[must_use]
    pub const fn created_styles(self) -> usize {
        self.created_styles
    }
}

/// One validated text-style publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Edit<'_> {
    /// Stage `style` over a UTF-16 `span` of the section's text.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidSpan`] for an empty span or one past the section;
    /// [`Refusal::EmptyStyle`] for a style that sets nothing.
    pub fn set(&mut self, span: TextSpan, style: CharacterProperties) -> Result<&mut Self, Error> {
        if style.is_empty() {
            return Err(Error::Refused(Refusal::EmptyStyle));
        }
        let range = section_range(self.source, self.section)?;
        if span.is_empty() || span.end().utf16_index() as usize > range.len() {
            return Err(Error::Refused(Refusal::InvalidSpan));
        }
        self.spans
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.spans.push((span, style));
        Ok(self)
    }

    /// Validate and atomically publish the staged styles.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::UnsupportedDependency`] for a body whose styles this
    /// transaction does not rewrite; source, resource and verification
    /// errors. Failure never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.spans.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let range = section_range(self.source, self.section)?;
        let mut absolute = Vec::new();
        absolute
            .try_reserve(self.spans.len())
            .map_err(|_error| Error::Allocation {
                amount: self.spans.len(),
            })?;
        for (span, style) in &self.spans {
            absolute.push((
                range.start + span.start().utf16_index() as usize
                    ..range.start + span.end().utf16_index() as usize,
                style,
            ));
        }
        let (candidate, created) = publish(self.source, &absolute)?;
        // Every staged span reads back with its properties set.
        let runs = section_runs(&candidate, self.section)?;
        for (span, style) in &self.spans {
            let covered = runs
                .iter()
                .filter(|run| run.span.start() < span.end() && run.span.end() > span.start());
            for run in covered {
                if !overrides_hold(&run.style, style) {
                    return Err(Error::Verification);
                }
            }
        }
        if candidate.sections().len() != self.source.sections().len()
            || candidate.sections()[self.section.get()].body_text()
                != self.source.sections()[self.section.get()].body_text()
        {
            return Err(Error::Verification);
        }
        let source_owner = SharedBytes::from_shared_slice(self.source.state.source.shared_source());
        let target_owner = SharedBytes::from_shared_slice(candidate.state.source.shared_source());
        Ok(Commit {
            package: candidate,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            },
            diagnostics: Diagnostics {
                created_styles: created,
            },
        })
    }
}

/// Whether `run` sets every property `requested` sets, to the same value.
/// A later staged style may override an earlier one on overlapping text,
/// so only the properties the run itself carries are compared.
fn overrides_hold(run: &CharacterProperties, requested: &CharacterProperties) -> bool {
    fn same<T>(run: Option<&T>, requested: Option<&T>) -> bool {
        requested.is_none() || run.is_some()
    }
    same(run.bold.as_ref(), requested.bold.as_ref())
        && same(run.italic.as_ref(), requested.italic.as_ref())
        && same(run.underline.as_ref(), requested.underline.as_ref())
        && same(run.strikethrough.as_ref(), requested.strikethrough.as_ref())
        && same(run.size.as_ref(), requested.size.as_ref())
        && same(run.font_name.as_ref(), requested.font_name.as_ref())
        && same(run.color.as_ref(), requested.color.as_ref())
}

fn section_range(package: &Package, section: Position) -> Result<std::ops::Range<usize>, Error> {
    let body = native_body(package).map_err(|_error| Error::InvalidSource)?;
    native_section_range(package, &body, section).map_err(|error| match error {
        super::section_text::Error::Refused(crate::section::text::Refusal::PositionNotFound {
            ..
        }) => Error::Refused(Refusal::SectionNotFound),
        _ => Error::InvalidSource,
    })
}

fn body_storage(package: &Package) -> Result<(u64, &ArchiveObject, usize), Error> {
    let body = native_body(package).map_err(|_error| Error::InvalidSource)?;
    let identifier = body.identifier.get();
    let object =
        find_object(package.state.source.components(), identifier).ok_or(Error::InvalidSource)?;
    let index = unique_message(object, STORAGE_MESSAGE_TYPE)?;
    Ok((identifier, object, index))
}

fn unique_message(object: &ArchiveObject, message_type: u32) -> Result<usize, Error> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ == message_type && found.replace(index).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn map_style_wire(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: LimitKind::Wire(kind),
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn style_wire_limits(package: &Package) -> Result<litchi_iwa_common::WireLimits, Error> {
    let maximum = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|error| map_archive(&error))?
        .max_message_bytes();
    litchi_iwa_common::WireLimits::default()
        .with_input_bytes(maximum.min(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| {
            limits.with_output_bytes(maximum.min(litchi_iwa_common::WireLimits::MAX_OUTPUT_BYTES))
        })
        .map_err(map_style_wire)
}

fn reference_field(
    payload: &[u8],
    number: u32,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<u64>, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(payload, wire).map_err(map_style_wire)?;
    let mut found = None;
    for field in view.fields() {
        if field.number() != number {
            continue;
        }
        let inner = litchi_iwa_common::wire::WireView::parse(field.payload(), wire)
            .map_err(map_style_wire)?;
        for reference in inner.fields() {
            if reference.number() == 1 {
                let (value, _) =
                    litchi_iwa_common::varint::decode_varint_from_bytes(reference.payload())
                        .map_err(|_error| Error::InvalidSource)?;
                if found.replace(value).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    Ok(found)
}

/// The runs of section `section` and the properties their styles set.
fn section_runs(package: &Package, section: Position) -> Result<Box<[Run]>, Error> {
    let wire = style_wire_limits(package)?;
    let range = section_range(package, section)?;
    let (_, object, index) = body_storage(package)?;
    let runs = litchi_iwa_text_wire::char_style_runs(&object.messages[index].data, wire)?;
    let mut output = Vec::new();
    for (position, run) in runs.iter().enumerate() {
        let end = runs
            .get(position + 1)
            .map_or(usize::MAX, |next| next.start)
            .min(range.end);
        let start = run.start.max(range.start);
        if start >= end {
            continue;
        }
        let style = match run.style {
            Some(style) => effective_style(package, style)?,
            None => CharacterProperties::new(),
        };
        output
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        output.push(Run {
            span: TextSpan::from_utf16_indexes(start - range.start, end - range.start)
                .map_err(|_error| Error::InvalidSource)?,
            style,
        });
    }
    Ok(output.into_boxed_slice())
}

/// The longest parent chain a run's style is resolved through.
const MAX_STYLE_DEPTH: usize = 16;

/// The character properties `style` and its parent chain set.
fn effective_style(package: &Package, style: u64) -> Result<CharacterProperties, Error> {
    let wire = style_wire_limits(package)?;
    let mut properties = CharacterProperties::new();
    let mut current = Some(style);
    for _ in 0..MAX_STYLE_DEPTH {
        let Some(identifier) = current else { break };
        let Some(object) = find_object(package.state.source.components(), identifier) else {
            break;
        };
        let Ok(message) = unique_message(object, CHARACTER_STYLE_MESSAGE_TYPE) else {
            break;
        };
        let payload = &object.messages[message].data;
        properties.inherit(&litchi_iwa_text_wire::character_style_properties(
            payload, wire,
        )?);
        current = litchi_iwa_text_wire::character_style_parent(payload, wire)?;
    }
    Ok(properties)
}

/// The registry facts a style insertion needs.
#[derive(Default)]
struct Registry {
    components: Vec<(u64, String)>,
    uuids: BTreeSet<(u64, u64)>,
    unknown: bool,
}

impl PackageMetadataVisitor for Registry {
    fn visit_unknown_field(&mut self) -> Result<(), RewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(&mut self, component: ComponentDescriptor<'_>) -> Result<(), RewriteError> {
        if component.is_current() {
            self.components.push((
                component.identifier(),
                component.effective_locator().to_owned(),
            ));
        }
        Ok(())
    }

    fn visit_object_uuid(&mut self, binding: ObjectUuidDescriptor<'_>) -> Result<(), RewriteError> {
        let uuid = binding.uuid();
        self.uuids.insert((uuid.lower(), uuid.upper()));
        Ok(())
    }
}

fn metadata_options(payload: usize, components: usize, additions: usize) -> RewriteOptions {
    RewriteOptions::new(
        payload.max(1),
        payload.saturating_mul(2).saturating_add(4_096),
        payload
            .saturating_mul(4)
            .saturating_add(additions.saturating_mul(16))
            .max(64),
        payload
            .saturating_mul(256)
            .saturating_add(
                components
                    .saturating_mul(additions.max(1))
                    .saturating_mul(64),
            )
            .clamp(1, litchi_iwa_common::WireLimits::MAX_REWRITE_WORK),
        64,
        components.saturating_add(additions),
        payload.max(1),
        additions.max(1),
    )
}

/// The locator a registry component uses for member `name`
/// (`Index/Document.iwa` -> `Document`).
fn locator(name: &str) -> &str {
    let name = name.strip_prefix("Index/").unwrap_or(name);
    name.strip_suffix(".iwa").unwrap_or(name)
}

/// A deterministic, collision-checked UUID for a new style.
fn style_uuid(identifier: u64, taken: &mut BTreeSet<(u64, u64)>) -> UuidBits {
    let mut state = identifier ^ 0x7465_7874_7374_796c;
    loop {
        let lower = mix(state);
        let upper = mix(lower ^ 0x6c69_7463_6869_0309) | 1;
        if taken.insert((lower, upper)) {
            return UuidBits::new(lower, upper);
        }
        state = state.wrapping_add(0x9e37_79b9_7f4a_7c15);
    }
}

const fn mix(mut value: u64) -> u64 {
    value = value.wrapping_add(0x9e37_79b9_7f4a_7c15);
    value = (value ^ (value >> 30)).wrapping_mul(0xbf58_476d_1ce4_e5b9);
    value = (value ^ (value >> 27)).wrapping_mul(0x94d0_49bb_1331_11eb);
    value ^ (value >> 31)
}

/// Decompress and parse one member for editing.
pub(super) fn parse_member(package: &Package, name: &str) -> Result<Archive, Error> {
    let catalog = &package.state.source;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let limits = catalog.limits();
    let snappy = limits
        .snappy_limits()
        .map_err(|error| map_archive(&error))?;
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(|error| map_archive(&error))?;
    let stream = SnappyStream::decompress(entry.data(), snappy)?;
    let archive = Archive::parse(stream.as_bytes(), archive_limits)?;
    archive.validate_canonical_object_framing(stream.as_bytes(), archive_limits)?;
    Ok(archive)
}

pub(super) fn component_name_of(package: &Package, identifier: u64) -> Result<String, Error> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        if component.archive().object(identifier).is_some()
            && found.replace(component.name().to_owned()).is_some()
        {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
    }
    found.ok_or(Error::InvalidSource)
}

/// Stage, publish and reopen the styled candidate; returns the number of
/// styles created.
#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: styles, stylesheet, storage, registry"
)]
fn publish(
    source: &Package,
    spans: &[(std::ops::Range<usize>, &CharacterProperties)],
) -> Result<(Package, usize), Error> {
    let wire = style_wire_limits(source)?;
    let catalog = &source.state.source;
    let limits = catalog.limits();
    let snappy_limits = limits
        .snappy_limits()
        .map_err(|error| map_archive(&error))?;
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(|error| map_archive(&error))?;
    let (storage, storage_object, storage_index) = body_storage(source)?;
    let mut payload = storage_object.messages[storage_index].data.clone();
    let storage_refs_before = storage_object
        .archive_info
        .message_infos
        .get(storage_index)
        .ok_or(Error::InvalidSource)?
        .object_references
        .clone();
    let stylesheet = reference_field(&payload, 2, wire)?
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let storage_member = component_name_of(source, storage)?;
    let sheet_member = component_name_of(source, stylesheet)?;
    let text_len = native_body(source)
        .map_err(|_error| Error::InvalidSource)?
        .utf16_len;

    // Registry facts and the next identifier.
    let metadata_archive = parse_member(source, METADATA_ENTRY_NAME)?;
    let (metadata_object, metadata_index) = {
        let mut found = None;
        for object in &metadata_archive.objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ == METADATA_MESSAGE_TYPE
                    && found
                        .replace((
                            object.archive_info.identifier.ok_or(Error::InvalidSource)?,
                            index,
                        ))
                        .is_some()
                {
                    return Err(Error::InvalidSource);
                }
            }
        }
        found.ok_or(Error::InvalidSource)?
    };
    let metadata_payload = metadata_archive
        .object(metadata_object)
        .ok_or(Error::InvalidSource)?
        .messages[metadata_index]
        .data
        .clone();
    // The registry may name more components than the catalog holds.
    let components = catalog.components().len().max(metadata_payload.len());
    let mut registry = Registry::default();
    let inspection = inspect_package_metadata_with_visitor(
        &metadata_payload,
        metadata_options(metadata_payload.len(), components, 0),
        &mut registry,
    )
    .map_err(|error| map_rewrite(&error))?;
    if registry.unknown {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let current_last = inspection.last_object_identifier();
    let mut maximum = current_last;
    for component in catalog.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(object.archive_info.identifier.ok_or(Error::InvalidSource)?);
        }
    }
    let mut next = maximum.checked_add(1).ok_or(Error::Limit {
        kind: LimitKind::Identifiers,
        observed: u64::MAX,
        maximum: u64::MAX,
    })?;
    let component_id = |member: &str| -> Result<(u64, String), Error> {
        let wanted = locator(member);
        let mut found = None;
        for (identifier, locator) in &registry.components {
            if locator == wanted && found.replace((*identifier, locator.clone())).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        found.ok_or(Error::Refused(Refusal::UnsupportedSource))
    };
    let storage_component = component_id(&storage_member)?;
    let sheet_component = component_id(&sheet_member)?;

    let mut sheet_archive = parse_member(source, &sheet_member)?;
    let mut storage_archive = if storage_member == sheet_member {
        None
    } else {
        Some(parse_member(source, &storage_member)?)
    };
    let versions = sheet_archive
        .objects
        .iter()
        .find_map(|object| {
            object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
                .find(|(message, _)| message.type_ == CHARACTER_STYLE_MESSAGE_TYPE)
                .map(|(_, info)| info.versions.clone())
        })
        .unwrap_or_else(|| CHARACTER_STYLE_VERSION.to_vec().into());

    let mut created: Vec<u64> = Vec::new();
    let mut taken = registry.uuids.clone();
    let mut uuids: Vec<(u64, UuidBits)> = Vec::new();
    for (range, properties) in spans {
        let previous =
            litchi_iwa_text_wire::styles_in_span(&payload, text_len, range.clone(), wire)?;
        let mut mapping: Vec<(Option<u64>, u64)> = Vec::new();
        for parent in previous {
            if let Some(parent) = parent
                && component_name_of_staged(source, &sheet_archive, &sheet_member, parent)?
                    != sheet_member
            {
                return Err(Error::Refused(Refusal::UnsupportedDependency));
            }
            let identifier = next;
            next = next.checked_add(1).ok_or(Error::Limit {
                kind: LimitKind::Identifiers,
                observed: u64::MAX,
                maximum: u64::MAX,
            })?;
            let data = litchi_iwa_text_wire::character_style_variation(
                parent, stylesheet, properties, wire,
            )?;
            let mut object = ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: CHARACTER_STYLE_MESSAGE_TYPE,
                    data,
                }],
                archive_limits,
            )?;
            let info = object
                .archive_info
                .message_infos
                .first_mut()
                .ok_or(Error::InvalidSource)?;
            info.versions.clone_from(&versions);
            info.object_references = parent.into_iter().chain([stylesheet]).collect();
            sheet_archive.insert_object(object, archive_limits)?;
            // The stylesheet lists the new style.
            let sheet = sheet_archive
                .object_mut(stylesheet)
                .ok_or(Error::InvalidSource)?;
            let sheet_index = unique_message(sheet, STYLESHEET_MESSAGE_TYPE)?;
            let sheet_data = litchi_iwa_text_wire::stylesheet_with_style(
                &sheet.messages[sheet_index].data,
                identifier,
                parent,
                wire,
            )?;
            let before = sheet.archive_info.message_infos[sheet_index]
                .object_references
                .clone();
            let mut after = before.clone();
            after.push(identifier);
            // A native header whose field-local references the aggregate
            // list does not repeat cannot be transitioned; its payload is
            // then replaced with the header kept as it is.
            let declared = sheet.replace_message_transitioning_object_references_preserving_header(
                sheet_index,
                RawMessage {
                    type_: STYLESHEET_MESSAGE_TYPE,
                    data: sheet_data.clone(),
                },
                litchi_iwa_core::archive::ObjectReferenceTransition {
                    aggregate_before: &before,
                    aggregate_after: &after,
                    fields: &[],
                },
                archive_limits,
            );
            if declared.is_err() {
                sheet.replace_message_preserving_header(
                    sheet_index,
                    RawMessage {
                        type_: STYLESHEET_MESSAGE_TYPE,
                        data: sheet_data,
                    },
                    archive_limits,
                )?;
            }
            uuids.push((identifier, style_uuid(identifier, &mut taken)));
            created.push(identifier);
            mapping.push((parent, identifier));
        }
        payload = litchi_iwa_text_wire::restyle_span(
            &payload,
            text_len,
            range.clone(),
            |style| {
                mapping
                    .iter()
                    .find(|(parent, _)| *parent == style)
                    .map(|(_, created)| *created)
            },
            wire,
        )?;
    }
    // The storage's new reference list.
    let still: BTreeSet<u64> = litchi_iwa_text_wire::char_style_runs(&payload, wire)?
        .iter()
        .filter_map(|run| run.style)
        .collect();
    let old_styles: BTreeSet<u64> =
        litchi_iwa_text_wire::char_style_runs(&storage_object.messages[storage_index].data, wire)?
            .iter()
            .filter_map(|run| run.style)
            .collect();
    let mut storage_after: Vec<u64> = storage_refs_before
        .iter()
        .copied()
        .filter(|identifier| !old_styles.contains(identifier) || still.contains(identifier))
        .collect();
    for identifier in &created {
        if still.contains(identifier) && !storage_after.contains(identifier) {
            storage_after.push(*identifier);
        }
    }
    {
        let archive = storage_archive.as_mut().unwrap_or(&mut sheet_archive);
        let object = archive.object_mut(storage).ok_or(Error::InvalidSource)?;
        object.replace_message_transitioning_object_references_preserving_header(
            storage_index,
            RawMessage {
                type_: STORAGE_MESSAGE_TYPE,
                data: payload,
            },
            litchi_iwa_core::archive::ObjectReferenceTransition {
                aggregate_before: &storage_refs_before,
                aggregate_after: &storage_after,
                fields: &[],
            },
            archive_limits,
        )?;
    }

    // Registry: UUIDs of the new styles, the body component's references
    // to them, and the watermark.
    let sheet_selector = ComponentSelector::new(sheet_component.0, &sheet_component.1);
    let storage_selector = ComponentSelector::new(storage_component.0, &storage_component.1);
    let uuid_additions: Vec<ObjectUuidAddition<'_>> = uuids
        .iter()
        .map(|(identifier, uuid)| ObjectUuidAddition::new(sheet_selector, *identifier, *uuid))
        .collect();
    let external_additions: Vec<ExternalReferenceAddition<'_>> = if storage_component.0
        == sheet_component.0
    {
        Vec::new()
    } else {
        created
            .iter()
            .filter(|identifier| still.contains(identifier))
            .map(|identifier| {
                ExternalReferenceAddition::new(storage_selector, sheet_selector, *identifier, None)
            })
            .collect()
    };
    let last = next - 1;
    let metadata = rewrite_package_metadata(
        &metadata_payload,
        Batch::new(current_last, last, &uuid_additions, &external_additions),
        metadata_options(
            metadata_payload.len(),
            components,
            uuid_additions.len() + external_additions.len(),
        ),
    )
    .map_err(|error| map_rewrite(&error))?
    .into_bytes();
    let mut metadata_archive = metadata_archive;
    metadata_archive
        .object_mut(metadata_object)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            metadata_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: metadata,
            },
            archive_limits,
        )?;

    // Reassemble once and reopen.
    let mut encoded: Vec<(String, Vec<u8>)> = Vec::new();
    let mut encode = |name: &str, archive: &Archive| -> Result<(), Error> {
        let bytes = archive.to_bytes(archive_limits)?;
        encoded.push((
            name.to_owned(),
            SnappyStream::compress(&bytes, snappy_limits)?,
        ));
        Ok(())
    };
    encode(&sheet_member, &sheet_archive)?;
    if let Some(archive) = &storage_archive {
        encode(&storage_member, archive)?;
    }
    encode(METADATA_ENTRY_NAME, &metadata_archive)?;
    let edits: Vec<EntryEdit<'_>> = encoded
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let bytes = catalog
        .package()
        .reassemble_with_changes_to_bytes(&[], &edits, &[], limits)
        .map_err(|error| map_archive(&error))?;
    let candidate = Package::from_bytes_with_archive_limits(&bytes, limits)
        .map_err(|_error| Error::Verification)?;
    // Every member the transaction did not rewrite keeps its bytes.
    let after = candidate.state.source.package();
    if after.iter().count() != catalog.package().iter().count() {
        return Err(Error::Verification);
    }
    for (before, now) in catalog.package().iter().zip(after.iter()) {
        let rewritten = encoded.iter().any(|(name, _)| name == before.name());
        if before.name() != now.name() || (!rewritten && before.data() != now.data()) {
            return Err(Error::Verification);
        }
    }
    Ok((candidate, created.len()))
}

/// The member holding `identifier`, looking at the staged stylesheet
/// archive first (it may hold styles created by this edit).
fn component_name_of_staged(
    source: &Package,
    staged: &Archive,
    staged_member: &str,
    identifier: u64,
) -> Result<String, Error> {
    if staged.object(identifier).is_some() {
        return Ok(staged_member.to_owned());
    }
    component_name_of(source, identifier)
}

impl Package {
    /// The character-style runs of one section's text: each run's UTF-16
    /// span and the properties its style (with its parents) sets.
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`]; source errors.
    pub fn section_text_runs<'selector>(
        &self,
        section: impl Into<SectionSelector<'selector>>,
    ) -> Result<Box<[Run]>, Error> {
        section_runs(self, resolve(self, section.into())?)
    }

    /// Start an immutable character-style edit of one section's text.
    ///
    /// ```no_run
    /// use litchi_pages::{Package, TextSpan};
    /// use litchi_pages::section::text_style::CharacterProperties;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("report.pages")?;
    /// let mut bold = CharacterProperties::new();
    /// bold.bold = Some(true);
    /// let mut edit = package.edit_section_text_style(0)?;
    /// edit.set(TextSpan::from_utf16_indexes(0, 5)?, bold)?;
    /// let commit = edit.commit()?;
    /// let restored = commit
    ///     .package()
    ///     .apply_section_text_style(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`].
    pub fn edit_section_text_style<'selector>(
        &self,
        section: impl Into<SectionSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            section: resolve(self, section.into())?,
            spans: Vec::new(),
        })
    }

    /// Apply a text-style patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_section_text_style(&self, patch: &Patch) -> Result<Commit, Error> {
        let source_owner = SharedBytes::from_shared_slice(self.state.source.shared_source());
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            let target = Arc::<[u8]>::from(patch.artifacts.target_owner().as_ref());
            Package::from_bytes_with_archive_limits(&target, self.state.source.limits())
                .map_err(|_error| Error::Verification)?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics::default(),
        })
    }
}

fn resolve(package: &Package, selector: SectionSelector<'_>) -> Result<Position, Error> {
    let selected = package
        .document_ref()
        .select_section(selector)
        .map_err(|_error| Error::Refused(Refusal::SectionNotFound))?
        .ok_or(Error::Refused(Refusal::SectionNotFound))?;
    Ok(Position::new(selected.index()))
}

#[cfg(test)]
mod wire_error_tests {
    use super::{Error, LimitKind, Registry, RewriteOptions, map_rewrite};
    use litchi_iwa_protos::package_metadata_codec::inspect_package_metadata_with_visitor;

    #[test]
    fn registry_limits_keep_the_codec_resource_and_counts() {
        // A last-object identifier and an unknown varint field are valid
        // registry wire data. Both count toward the explicit wire policy.
        let source = [8, 1, 0xa0, 1, 7];
        let inspect = |input, fields, work| {
            inspect_package_metadata_with_visitor(
                &source,
                RewriteOptions::new(input, 100, fields, work, 64, 10, 10, 10),
                &mut Registry::default(),
            )
        };
        let report = inspect(source.len(), 100, 1_000)
            .expect("bounded registry inspection")
            .report();
        // The codec charges both its preflight and visitor passes.
        inspect(source.len(), report.fields(), report.work_bytes())
            .expect("exact aggregate limits");
        let input = inspect(source.len() - 1, report.fields(), report.work_bytes()).unwrap_err();
        assert!(matches!(
            map_rewrite(&input),
            Error::Limit {
                kind: LimitKind::RegistryBytes,
                observed: 5,
                maximum: 4,
            }
        ));
        let fields = inspect(source.len(), report.fields() - 1, report.work_bytes()).unwrap_err();
        assert!(matches!(
            map_rewrite(&fields),
            Error::Limit {
                kind: LimitKind::RegistryFields,
                observed,
                maximum,
            } if observed == report.fields() as u64 && maximum == (report.fields() - 1) as u64
        ));
        let work = inspect(source.len(), report.fields(), report.work_bytes() - 1).unwrap_err();
        assert!(matches!(
            map_rewrite(&work),
            Error::Limit {
                kind: LimitKind::RegistryWork,
                observed,
                maximum,
            } if observed == report.work_bytes() as u64 && maximum == (report.work_bytes() - 1) as u64
        ));
        let malformed = inspect_package_metadata_with_visitor(
            &[0],
            RewriteOptions::new(10, 100, 10, 1_000, 64, 10, 10, 10),
            &mut Registry::default(),
        )
        .unwrap_err();
        assert!(matches!(map_rewrite(&malformed), Error::InvalidSource));
    }

    #[test]
    fn character_style_resource_errors_survive_the_owner_boundary() {
        let error = Error::from(litchi_iwa_text_wire::CharStyleError::Limit {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed: 11,
            maximum: 10,
        });
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::Wire(litchi_iwa_common::LimitKind::RewriteWork),
                observed: 11,
                maximum: 10,
            }
        ));
        assert!(matches!(
            Error::from(litchi_iwa_text_wire::CharStyleError::Allocation { amount: 42 }),
            Error::Allocation { amount: 42 }
        ));
        assert!(matches!(
            Error::from(litchi_iwa_text_wire::CharStyleError::Invalid),
            Error::InvalidSource
        ));
    }
}
