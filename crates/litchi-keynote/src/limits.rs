//! Caller-selected ingress limits and the Keynote open-options envelope.
//!
//! Every Keynote ingress route — [`crate::Package`] and the archive-free
//! [`crate::Presentation`] projection — accepts one [`OpenOptions`] value.
//! The envelope carries checked physical [`SourceLimits`] (input bytes, ZIP
//! catalog dimensions, decoded IWA components and their archive records) and
//! checked [`SemanticLimits`] (show-graph dimensions enforced while the
//! semantic projection is decoded).
//!
//! Construction follows the workspace idiom: a consuming builder with fallible
//! `max_<dimension>` setters and a fallible `build()`. A value outside
//! `1..=maximum` is reported as [`crate::Error::InvalidLimit`].

use std::fmt;

use litchi_iwa_core::Limits as IwaArchiveLimits;

/// Hard ceiling for native IWA objects indexed by one Keynote package.
pub(crate) const MAX_OBJECTS: usize = 1_000_000;
/// Hard ceiling for semantic slides decoded from one Keynote package.
pub(crate) const MAX_SLIDES: usize = 65_536;
/// Hard ceiling for semantic graph-reference occurrences traversed by one package.
pub(crate) const MAX_REFERENCES: usize = 1_000_000;
/// Hard ceiling for text-storage objects decoded from one Keynote package.
pub(crate) const MAX_TEXT_STORAGES: usize = 1_000_000;
/// Hard ceiling for rich-text fragment ranges retained by one Keynote package.
pub(crate) const MAX_TEXT_FRAGMENTS: usize = 1_000_000;
/// Hard ceiling for aggregate decoded text bytes in one Keynote package.
pub(crate) const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

fn usize_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn u64_usize(value: u64) -> usize {
    usize::try_from(value).unwrap_or(usize::MAX)
}

/// The physical ingress dimension named by a source limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Bytes of the complete physical source (file or buffer).
    InputBytes,
    /// Physical ZIP members, including directory records.
    Entries,
    /// Declared uncompressed bytes of one ZIP member.
    EntryBytes,
    /// Aggregate declared uncompressed bytes of all ZIP members.
    AggregateBytes,
    /// Decompressed bytes of one IWA component.
    ComponentBytes,
    /// Stored objects in one decompressed IWA component.
    ComponentObjects,
    /// Stored messages in one decompressed IWA component.
    ComponentMessages,
    /// Payload bytes of one stored message.
    MessageBytes,
    /// Fields inspected across one object header.
    HeaderFields,
}

impl SourceLimitKind {
    /// Every source dimension, in declaration order.
    pub const ALL: [Self; 9] = [
        Self::InputBytes,
        Self::Entries,
        Self::EntryBytes,
        Self::AggregateBytes,
        Self::ComponentBytes,
        Self::ComponentObjects,
        Self::ComponentMessages,
        Self::MessageBytes,
        Self::HeaderFields,
    ];

    /// The registered glossary noun of this dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::InputBytes => "input_bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry_bytes",
            Self::AggregateBytes => "aggregate_bytes",
            Self::ComponentBytes => "component_bytes",
            Self::ComponentObjects => "component_objects",
            Self::ComponentMessages => "component_messages",
            Self::MessageBytes => "message_bytes",
            Self::HeaderFields => "header_fields",
        }
    }
}

impl fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Checked physical limits for one Keynote ingress.
///
/// `max_input_bytes` bounds the physical source before ZIP indexing starts;
/// the entry and aggregate dimensions bound the ZIP catalog; the component
/// dimension bounds each decompressed IWA stream, and the archive dimensions
/// bound the object, message and header records parsed from one component.
/// The effective per-message ceiling is the smaller of
/// `max_message_bytes` and `max_component_bytes`.
///
/// The defaults equal the hard ceilings of the physical iWork archive reader
/// ([`SourceLimits::MAXIMUM`]), so a default open behaves exactly as before
/// this envelope existed. The archive-free [`crate::Presentation`] ingress
/// additionally caps the input at the format detector's own hard ceiling.
///
/// ```
/// use litchi_keynote::{OpenOptions, SourceLimits};
///
/// # fn main() -> Result<(), litchi_keynote::Error> {
/// let limits = SourceLimits::builder()
///     .max_input_bytes(32 * 1024 * 1024)?
///     .max_entries(4_096)?
///     .build()?;
/// let options = OpenOptions::new().source_limits(limits);
/// assert_eq!(options.source_limits_value().max_entries(), 4_096);
/// # Ok(())
/// # }
/// ```
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: u64,
    entries: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
    component_objects: usize,
    component_messages: usize,
    message_bytes: usize,
    header_fields: usize,
}

impl SourceLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        input_bytes: litchi_iwa_archive::Limits::MAX_INPUT_BYTES,
        entries: litchi_iwa_archive::Limits::MAX_ENTRIES,
        entry_bytes: litchi_iwa_archive::Limits::MAX_ENTRY_BYTES,
        aggregate_bytes: litchi_iwa_archive::Limits::MAX_TOTAL_BYTES,
        component_bytes: litchi_iwa_archive::Limits::MAX_IWA_STREAM_BYTES,
        component_objects: IwaArchiveLimits::MAX_OBJECTS,
        component_messages: IwaArchiveLimits::MAX_MESSAGES,
        message_bytes: IwaArchiveLimits::MAX_MESSAGE_BYTES,
        header_fields: IwaArchiveLimits::MAX_HEADER_FIELDS,
    };

    /// The default ingress profile, equal to [`Self::MAXIMUM`].
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum bytes of the complete physical source.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum physical ZIP members, including directory records.
    #[must_use]
    pub const fn max_entries(self) -> usize {
        self.entries
    }

    /// Maximum declared uncompressed bytes of one ZIP member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    /// Maximum aggregate declared uncompressed bytes.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    /// Maximum decompressed bytes of one IWA component.
    #[must_use]
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    /// Maximum archive objects in one IWA component.
    #[must_use]
    pub const fn max_component_objects(self) -> usize {
        self.component_objects
    }

    /// Maximum archive messages in one IWA component.
    #[must_use]
    pub const fn max_component_messages(self) -> usize {
        self.component_messages
    }

    /// Maximum payload bytes of one archive message.
    #[must_use]
    pub const fn max_message_bytes(self) -> usize {
        self.message_bytes
    }

    /// Maximum fields inspected across one archive header wire tree.
    #[must_use]
    pub const fn max_header_fields(self) -> usize {
        self.header_fields
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SourceLimitKind) -> u64 {
        match kind {
            SourceLimitKind::InputBytes => self.input_bytes,
            SourceLimitKind::Entries => self.entries as u64,
            SourceLimitKind::EntryBytes => self.entry_bytes,
            SourceLimitKind::AggregateBytes => self.aggregate_bytes,
            SourceLimitKind::ComponentBytes => self.component_bytes as u64,
            SourceLimitKind::ComponentObjects => self.component_objects as u64,
            SourceLimitKind::ComponentMessages => self.component_messages as u64,
            SourceLimitKind::MessageBytes => self.message_bytes as u64,
            SourceLimitKind::HeaderFields => self.header_fields as u64,
        }
    }

    /// The physical iWork archive profile used by package ingress.
    pub(crate) fn archive_limits(self) -> litchi_iwa_archive::Result<litchi_iwa_archive::Limits> {
        let invalid = |error: litchi_iwa_core::Error| {
            litchi_iwa_archive::Error::InvalidLimits(error.to_string())
        };
        let limits = litchi_iwa_archive::Limits::new(
            self.input_bytes,
            self.entries,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )?;
        let profile = limits
            .archive_limits()
            .with_objects(self.component_objects)
            .map_err(invalid)?
            .with_messages(self.component_messages)
            .map_err(invalid)?
            .with_message_bytes(self.message_bytes.min(self.component_bytes))
            .map_err(invalid)?
            .with_header_fields(self.header_fields)
            .map_err(invalid)?;
        limits.with_archive_limits(profile)
    }

    /// The format-detection profile used by archive-free projection ingress.
    ///
    /// The detector has its own, lower hard input ceiling; the input
    /// dimension is clamped to it rather than rejected.
    pub(crate) fn detector_limits(self) -> litchi_iwa_detect::Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.input_bytes
                .min(litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES),
            self.entries,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
    }

    /// Describe an already validated physical archive profile. Values are
    /// clamped to the hard ceilings, never rejected.
    pub(crate) fn from_archive(limits: litchi_iwa_archive::Limits) -> Self {
        let maximum = Self::MAXIMUM;
        let profile = limits.archive_limits();
        Self {
            input_bytes: limits.max_input_bytes().clamp(1, maximum.input_bytes),
            entries: limits.max_entries().clamp(1, maximum.entries),
            entry_bytes: limits.max_entry_bytes().clamp(1, maximum.entry_bytes),
            aggregate_bytes: limits.max_total_bytes().clamp(1, maximum.aggregate_bytes),
            component_bytes: limits
                .max_iwa_stream_bytes()
                .clamp(1, maximum.component_bytes),
            component_objects: profile.max_objects().clamp(1, maximum.component_objects),
            component_messages: profile.max_messages().clamp(1, maximum.component_messages),
            message_bytes: profile.max_message_bytes().clamp(1, maximum.message_bytes),
            header_fields: profile.max_header_fields().clamp(1, maximum.header_fields),
        }
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SourceLimits`].
#[derive(Clone, Copy, Debug)]
#[must_use]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

fn checked(kind: SourceLimitKind, value: u64) -> Result<u64, crate::Error> {
    let maximum = SourceLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(crate::Error::InvalidLimit {
            kind: LimitKind::Source(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

fn checked_usize(kind: SourceLimitKind, value: usize) -> Result<usize, crate::Error> {
    // The maximum of every `usize` dimension is a `usize` value, so a
    // checked value fits.
    checked(kind, usize_u64(value)).map(u64_usize)
}

impl SourceLimitsBuilder {
    /// Set the maximum physical source bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_input_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.input_bytes = checked(SourceLimitKind::InputBytes, value)?;
        Ok(self)
    }

    /// Set the maximum physical ZIP members.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_entries(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.entries = checked_usize(SourceLimitKind::Entries, value)?;
        Ok(self)
    }

    /// Set the maximum declared uncompressed bytes of one ZIP member.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_entry_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.entry_bytes = checked(SourceLimitKind::EntryBytes, value)?;
        Ok(self)
    }

    /// Set the maximum aggregate declared uncompressed bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_aggregate_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.aggregate_bytes = checked(SourceLimitKind::AggregateBytes, value)?;
        Ok(self)
    }

    /// Set the maximum decompressed bytes of one IWA component.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_component_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.component_bytes = checked_usize(SourceLimitKind::ComponentBytes, value)?;
        Ok(self)
    }

    /// Set the maximum archive objects in one IWA component.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_component_objects(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.component_objects = checked_usize(SourceLimitKind::ComponentObjects, value)?;
        Ok(self)
    }

    /// Set the maximum archive messages in one IWA component.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_component_messages(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.component_messages = checked_usize(SourceLimitKind::ComponentMessages, value)?;
        Ok(self)
    }

    /// Set the maximum payload bytes of one archive message.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_message_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.message_bytes = checked_usize(SourceLimitKind::MessageBytes, value)?;
        Ok(self)
    }

    /// Set the maximum fields inspected across one archive header wire tree.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_header_fields(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.header_fields = checked_usize(SourceLimitKind::HeaderFields, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Every setter already validated its dimension; this re-validates the
    /// complete profile so a future cross-dimension rule has one home.
    pub fn build(self) -> Result<SourceLimits, crate::Error> {
        for kind in SourceLimitKind::ALL {
            checked(kind, self.limits.value(kind))?;
        }
        Ok(self.limits)
    }
}

/// The semantic projection dimension named by a semantic limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Native IWA objects indexed for package-wide lookup.
    Objects,
    /// Slides decoded from the native show tree.
    Slides,
    /// Semantic graph-reference occurrences traversed from the show root.
    References,
    /// Native text-storage objects decoded by the semantic reader.
    TextStorages,
    /// Rich-text fragment ranges retained by the semantic reader.
    TextFragments,
    /// Aggregate bytes retained from text storage and semantic identifiers.
    TextBytes,
}

impl SemanticLimitKind {
    /// Every semantic dimension, in declaration order.
    pub const ALL: [Self; 6] = [
        Self::Objects,
        Self::Slides,
        Self::References,
        Self::TextStorages,
        Self::TextFragments,
        Self::TextBytes,
    ];

    /// The registered glossary noun of this dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Objects => "objects",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text_storages",
            Self::TextFragments => "text_fragments",
            Self::TextBytes => "text_bytes",
        }
    }
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Checked resource ceilings for semantic decoding of one Keynote package.
///
/// The object index is bounded when a package is opened; the remaining
/// dimensions are enforced lazily on first semantic access, or eagerly by the
/// archive-free [`crate::Presentation`] ingress. The defaults equal the hard
/// ceilings ([`SemanticLimits::MAXIMUM`]).
///
/// ```
/// use litchi_keynote::{OpenOptions, SemanticLimits};
///
/// # fn main() -> Result<(), litchi_keynote::Error> {
/// let limits = SemanticLimits::builder().max_slides(512)?.build()?;
/// let options = OpenOptions::new().semantic_limits(limits);
/// assert_eq!(options.semantic_limits_value().max_slides(), 512);
/// # Ok(())
/// # }
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    objects: usize,
    slides: usize,
    references: usize,
    text_storages: usize,
    text_fragments: usize,
    text_bytes: usize,
}

impl SemanticLimits {
    /// Hard ceiling for native IWA objects indexed by one Keynote package.
    pub const MAX_OBJECTS: usize = MAX_OBJECTS;
    /// Hard ceiling for semantic slides decoded from one Keynote package.
    pub const MAX_SLIDES: usize = MAX_SLIDES;
    /// Hard ceiling for semantic graph-reference occurrences.
    pub const MAX_REFERENCES: usize = MAX_REFERENCES;
    /// Hard ceiling for decoded text-storage objects in one Keynote package.
    pub const MAX_TEXT_STORAGES: usize = MAX_TEXT_STORAGES;
    /// Hard ceiling for retained rich-text fragment ranges.
    pub const MAX_TEXT_FRAGMENTS: usize = MAX_TEXT_FRAGMENTS;
    /// Hard ceiling for aggregate decoded text bytes in one Keynote package.
    pub const MAX_TEXT_BYTES: usize = MAX_TEXT_BYTES;

    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        objects: MAX_OBJECTS,
        slides: MAX_SLIDES,
        references: MAX_REFERENCES,
        text_storages: MAX_TEXT_STORAGES,
        text_fragments: MAX_TEXT_FRAGMENTS,
        text_bytes: MAX_TEXT_BYTES,
    };

    /// The default profile, equal to [`Self::MAXIMUM`].
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum number of native IWA objects indexed for package-wide lookup.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.objects
    }

    /// Maximum number of semantic slides decoded from the native show tree.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.slides
    }

    /// Maximum semantic graph-reference occurrences traversed.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.references
    }

    /// Maximum number of native text-storage objects decoded.
    #[must_use]
    pub const fn max_text_storages(self) -> usize {
        self.text_storages
    }

    /// Maximum rich-text fragment ranges retained.
    #[must_use]
    pub const fn max_text_fragments(self) -> usize {
        self.text_fragments
    }

    /// Maximum aggregate byte length of retained semantic text and identifiers.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SemanticLimitKind) -> u64 {
        (match kind {
            SemanticLimitKind::Objects => self.objects,
            SemanticLimitKind::Slides => self.slides,
            SemanticLimitKind::References => self.references,
            SemanticLimitKind::TextStorages => self.text_storages,
            SemanticLimitKind::TextFragments => self.text_fragments,
            SemanticLimitKind::TextBytes => self.text_bytes,
        }) as u64
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SemanticLimits`].
#[derive(Clone, Copy, Debug)]
#[must_use]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}

fn checked_semantic(kind: SemanticLimitKind, value: usize) -> Result<usize, crate::Error> {
    let maximum = SemanticLimits::MAXIMUM.value(kind);
    let requested = usize_u64(value);
    if value == 0 || requested > maximum {
        return Err(crate::Error::InvalidLimit {
            kind: LimitKind::Semantic(kind),
            value: requested,
            maximum,
        });
    }
    Ok(value)
}

impl SemanticLimitsBuilder {
    /// Set the maximum native IWA objects indexed for package-wide lookup.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_objects(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.objects = checked_semantic(SemanticLimitKind::Objects, value)?;
        Ok(self)
    }

    /// Set the maximum semantic slides.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_slides(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.slides = checked_semantic(SemanticLimitKind::Slides, value)?;
        Ok(self)
    }

    /// Set the maximum semantic graph-reference occurrences.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_references(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.references = checked_semantic(SemanticLimitKind::References, value)?;
        Ok(self)
    }

    /// Set the maximum native text-storage objects decoded.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_text_storages(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.text_storages = checked_semantic(SemanticLimitKind::TextStorages, value)?;
        Ok(self)
    }

    /// Set the maximum rich-text fragment ranges retained.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_text_fragments(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.text_fragments = checked_semantic(SemanticLimitKind::TextFragments, value)?;
        Ok(self)
    }

    /// Set the maximum aggregate retained semantic text bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_text_bytes(mut self, value: usize) -> Result<Self, crate::Error> {
        self.limits.text_bytes = checked_semantic(SemanticLimitKind::TextBytes, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Re-validates every dimension.
    pub fn build(self) -> Result<SemanticLimits, crate::Error> {
        for kind in SemanticLimitKind::ALL {
            checked_semantic(kind, u64_usize(self.limits.value(kind)))?;
        }
        Ok(self.limits)
    }
}

/// The dimension named by a Keynote limit error.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// A physical ingress dimension.
    Source(SourceLimitKind),
    /// A semantic projection dimension.
    Semantic(SemanticLimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Source(kind) => kind.fmt(formatter),
            Self::Semantic(kind) => kind.fmt(formatter),
        }
    }
}

impl From<SourceLimitKind> for LimitKind {
    fn from(kind: SourceLimitKind) -> Self {
        Self::Source(kind)
    }
}

impl From<SemanticLimitKind> for LimitKind {
    fn from(kind: SemanticLimitKind) -> Self {
        Self::Semantic(kind)
    }
}

/// Configuration for one Keynote ingress.
///
/// Setters consume already validated policy values and perform no I/O.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[must_use]
pub struct OpenOptions {
    source_limits: SourceLimits,
    semantic_limits: SemanticLimits,
}

impl OpenOptions {
    /// The default ingress policies.
    pub const fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            semantic_limits: SemanticLimits::DEFAULT,
        }
    }

    /// Select the physical source limits.
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select the semantic projection limits.
    pub const fn semantic_limits(mut self, limits: SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The selected physical source limits.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// The selected semantic projection limits.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> SemanticLimits {
        self.semantic_limits
    }
}

#[cfg(test)]
mod tests {
    use super::{
        LimitKind, MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS,
        MAX_TEXT_STORAGES, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind,
        SourceLimits,
    };
    use crate::Error;

    #[test]
    fn defaults_use_every_hard_ceiling() {
        let limits = SemanticLimits::default();
        assert_eq!(limits.max_objects(), MAX_OBJECTS);
        assert_eq!(limits.max_slides(), MAX_SLIDES);
        assert_eq!(limits.max_references(), MAX_REFERENCES);
        assert_eq!(limits.max_text_storages(), MAX_TEXT_STORAGES);
        assert_eq!(limits.max_text_fragments(), MAX_TEXT_FRAGMENTS);
        assert_eq!(limits.max_text_bytes(), MAX_TEXT_BYTES);
        assert_eq!(SourceLimits::default(), SourceLimits::MAXIMUM);
        assert_eq!(OpenOptions::new(), OpenOptions::default());
    }

    #[test]
    fn default_source_profile_preserves_the_archive_defaults() {
        let archive = SourceLimits::default()
            .archive_limits()
            .expect("default source limits convert");
        assert_eq!(archive, litchi_iwa_archive::Limits::default());
        assert_eq!(SourceLimits::from_archive(archive), SourceLimits::default());
        let detector = SourceLimits::default()
            .detector_limits()
            .expect("default source limits convert");
        assert_eq!(detector, litchi_iwa_detect::Limits::default());
    }

    #[test]
    fn builders_retain_each_limit() {
        let limits = SemanticLimits::builder()
            .max_objects(1)
            .and_then(|builder| builder.max_slides(2))
            .and_then(|builder| builder.max_references(3))
            .and_then(|builder| builder.max_text_storages(4))
            .and_then(|builder| builder.max_text_fragments(5))
            .and_then(|builder| builder.max_text_bytes(6))
            .and_then(super::SemanticLimitsBuilder::build)
            .expect("positive limits below the hard ceilings are valid");
        assert_eq!(limits.max_objects(), 1);
        assert_eq!(limits.max_slides(), 2);
        assert_eq!(limits.max_references(), 3);
        assert_eq!(limits.max_text_storages(), 4);
        assert_eq!(limits.max_text_fragments(), 5);
        assert_eq!(limits.max_text_bytes(), 6);

        let source = SourceLimits::builder()
            .max_component_bytes(1_024)
            .and_then(|builder| builder.max_header_fields(7))
            .and_then(super::SourceLimitsBuilder::build)
            .expect("positive limits below the hard ceilings are valid");
        let archive = source.archive_limits().expect("source limits convert");
        assert_eq!(archive.max_iwa_stream_bytes(), 1_024);
        assert_eq!(archive.archive_limits().max_message_bytes(), 1_024);
        assert_eq!(archive.archive_limits().max_header_fields(), 7);
    }

    #[test]
    fn builders_reject_zero_and_values_above_the_hard_ceiling() {
        let zero = SourceLimits::builder().max_input_bytes(0).unwrap_err();
        assert!(matches!(
            zero,
            Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                value: 0,
                ..
            }
        ));
        let above = SemanticLimits::MAX_SLIDES + 1;
        let error = SemanticLimits::builder().max_slides(above).unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: LimitKind::Semantic(SemanticLimitKind::Slides),
                value,
                maximum,
            } if value == above as u64 && maximum == SemanticLimits::MAX_SLIDES as u64
        ));
        for kind in SourceLimitKind::ALL {
            assert_eq!(kind.to_string(), kind.as_str());
        }
    }
}
