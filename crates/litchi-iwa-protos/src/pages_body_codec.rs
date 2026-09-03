//! Strict private Buffa projection for the Pages root/body graph leaves.
//!
//! A bounded raw-wire pass rejects ambiguous encodings before Buffa observes
//! the payload. Buffa then supplies a borrowed lazy-view cross-check for the
//! root body references and one section-boundary entry. Generated types never
//! cross this module, and caller-owned source bytes remain the sole
//! preservation representation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Semantic preflight intentionally precedes the low-level wire reader."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_pages_body_generated::LitchiIwaProjection as projection;

const DOCUMENT_BODY_STORAGE_FIELD: u32 = 4;
const DOCUMENT_INITIAL_SECTION_FIELD: u32 = 5;
const DOCUMENT_SUPER_FIELD: u32 = 15;
const DOCUMENT_FLOATING_DRAWABLES_FIELD: u32 = 3;
const DOCUMENT_THEME_FIELD: u32 = 6;
const DOCUMENT_DRAWABLES_ZORDER_FIELD: u32 = 20;
const DOCUMENT_LEFT_MARGIN_FIELD: u32 = 32;
const DOCUMENT_PAGE_TEMPLATES_FIELD: u32 = 48;
const DOCUMENT_SUPER_SUPER_FIELD: u32 = 1;
const DOCUMENT_LANGUAGE_FIELD: u32 = 3;
const BOUNDARY_CHARACTER_INDEX_FIELD: u32 = 1;
const BOUNDARY_SECTION_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Explicit finite resource policy for one focused Pages projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build a finite bytes/fields/work/nesting policy.
    ///
    /// Work counts a conservative strict-plus-Buffa scan of each projected
    /// message. A document containing references therefore charges the root
    /// bytes twice and each selected reference payload twice.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_fields: usize,
        max_work_bytes: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_fields,
            max_work_bytes,
            recursion_limit,
        }
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(self.max_fields)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self
            .recursion_limit
            .checked_sub(1)
            .ok_or_else(DecodeError::recursion_limit)?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Generated-free, strictly validated `TSP.Reference` values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReferenceSnapshot {
    identifier: NonZeroU64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

impl ReferenceSnapshot {
    /// Non-zero native object identifier.
    #[must_use]
    pub const fn identifier(self) -> NonZeroU64 {
        self.identifier
    }

    /// Optional deprecated native object-type hint.
    #[must_use]
    pub const fn deprecated_type(self) -> Option<i32> {
        self.deprecated_type
    }

    /// Optional deprecated external-reference flag.
    #[must_use]
    pub const fn deprecated_is_external(self) -> Option<bool> {
        self.deprecated_is_external
    }
}

/// Borrow-free focused projection of `TP.DocumentArchive`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DocumentBodySnapshot {
    body_storage: Option<ReferenceSnapshot>,
    initial_section: Option<ReferenceSnapshot>,
}

impl DocumentBodySnapshot {
    /// Optional body text-storage reference from field 4.
    #[must_use]
    pub const fn body_storage(self) -> Option<ReferenceSnapshot> {
        self.body_storage
    }

    /// Optional initial section reference from field 5.
    #[must_use]
    pub const fn initial_section(self) -> Option<ReferenceSnapshot> {
        self.initial_section
    }
}

/// Borrowed focused projection of the Pages document-root facts used by the
/// migration host.
//
// The root payload itself is retained as a borrow so callers can keep the
// native bytes authoritative. Singular references and the page-layout scalar
// are copied value facts; repeated page-template references are streamed with
// [`Self::page_templates`] and are never collected into an input-sized
// generated Buffa view.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct DocumentRootSnapshot<'source> {
    source: &'source [u8],
    options: DecodeOptions,
    floating_drawables: Option<ReferenceSnapshot>,
    theme: Option<ReferenceSnapshot>,
    drawables_zorder: Option<ReferenceSnapshot>,
    page_template_count: usize,
    left_margin: Option<f32>,
    document_language: Option<&'source str>,
}

impl<'source> DocumentRootSnapshot<'source> {
    /// Return the exact caller-owned root payload.
    #[must_use]
    pub const fn raw(self) -> &'source [u8] {
        self.source
    }

    /// Optional `TP.DocumentArchive.floating_drawables` reference.
    #[must_use]
    pub const fn floating_drawables(self) -> Option<ReferenceSnapshot> {
        self.floating_drawables
    }

    /// Optional `TP.DocumentArchive.theme` reference.
    #[must_use]
    pub const fn theme(self) -> Option<ReferenceSnapshot> {
        self.theme
    }

    /// Optional `TP.DocumentArchive.drawables_zorder` reference.
    #[must_use]
    pub const fn drawables_zorder(self) -> Option<ReferenceSnapshot> {
        self.drawables_zorder
    }

    /// Number of repeated `TP.DocumentArchive.page_templates` references.
    #[must_use]
    pub const fn page_template_count(self) -> usize {
        self.page_template_count
    }

    /// Optional `TP.DocumentArchive.left_margin` value.
    #[must_use]
    pub const fn left_margin(self) -> Option<f32> {
        self.left_margin
    }

    /// Optional `TSA.DocumentArchive.document_language` value.
    #[must_use]
    pub const fn document_language(self) -> Option<&'source str> {
        self.document_language
    }

    /// Stream page-template references in native source order.
    //
    // Each item is independently checked against the private lazy Buffa
    // `TSP.Reference` view. The strict root pass performed by
    // [`decode_document_root`] has already validated every item, so this
    // iterator does not retain a collection or defer an unchecked payload.
    #[must_use]
    pub fn page_templates(self) -> PageTemplateIter<'source> {
        PageTemplateIter {
            remaining: self.source,
            yielded: 0,
            maximum: self.page_template_count,
            options: self.options,
        }
    }
}

/// Allocation-free iterator over repeated Pages page-template references.
#[derive(Debug, Clone, Copy)]
pub struct PageTemplateIter<'source> {
    remaining: &'source [u8],
    yielded: usize,
    maximum: usize,
    options: DecodeOptions,
}

impl<'source> Iterator for PageTemplateIter<'source> {
    type Item = Result<ReferenceSnapshot, DecodeError>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.yielded >= self.maximum {
            return None;
        }
        let mut budget = Budget::new(self.options);
        loop {
            let field = match next_strict_field(&mut self.remaining, self.options, &mut budget) {
                Ok(Some(field)) => field,
                Ok(None) => {
                    self.yielded = self.maximum;
                    return Some(Err(DecodeError::projection()));
                },
                Err(error) => {
                    self.remaining = &[];
                    self.yielded = self.maximum;
                    return Some(Err(error));
                },
            };
            if field.number != DOCUMENT_PAGE_TEMPLATES_FIELD {
                continue;
            }
            let payload = match field.length_delimited() {
                Ok(payload) => payload,
                Err(error) => {
                    self.remaining = &[];
                    self.yielded = self.maximum;
                    return Some(Err(error));
                },
            };
            let nested_options = match self.options.descend() {
                Ok(options) => options,
                Err(error) => {
                    self.remaining = &[];
                    self.yielded = self.maximum;
                    return Some(Err(error));
                },
            };
            let strict = match preflight_reference(payload, nested_options, &mut budget) {
                Ok(strict) => strict,
                Err(error) => {
                    self.remaining = &[];
                    self.yielded = self.maximum;
                    return Some(Err(error));
                },
            };
            self.yielded = self.yielded.saturating_add(1);
            return Some(cross_check_reference(payload, strict, self.options));
        }
    }
}

impl std::iter::FusedIterator for PageTemplateIter<'_> {}

/// Borrow-free focused projection of one Pages section-boundary entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SectionBoundarySnapshot {
    character_index: u32,
    section: Option<ReferenceSnapshot>,
}

impl SectionBoundarySnapshot {
    /// UTF-16 character index stored by the native entry.
    #[must_use]
    pub const fn character_index(self) -> u32 {
        self.character_index
    }

    /// Optional native section reference stored in the entry's `object` field.
    #[must_use]
    pub const fn section(self) -> Option<ReferenceSnapshot> {
        self.section
    }
}

/// Failure from strict Pages body-graph preflight or its Buffa cross-check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    ZeroIdentifier(&'static str),
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Projection,
}

impl DecodeError {
    fn recursion_limit() -> Self {
        buffa::DecodeError::RecursionLimitExceeded.into()
    }

    const fn missing_required(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::MissingRequired(field),
        }
    }

    const fn duplicate_singular(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::DuplicateSingular(field),
        }
    }

    const fn noncanonical(reason: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::NonCanonical(reason),
        }
    }

    const fn zero_identifier(field: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::ZeroIdentifier(field),
        }
    }

    const fn field_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::FieldLimit { observed, maximum },
        }
    }

    const fn work_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::WorkLimit { observed, maximum },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Required known field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Singular known field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Stable canonicality failure reason, when applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Reference field carrying a forbidden zero identifier, when applicable.
    #[must_use]
    pub const fn zero_identifier_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::ZeroIdentifier(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::FieldLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::WorkLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::ZeroIdentifier(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::ZeroIdentifier(field) => {
                write!(formatter, "{field} is zero")
            },
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Pages body projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Pages body projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Projection => formatter
                .write_str("Pages body strict preflight disagrees with the Buffa projection"),
        }
    }
}

impl std::error::Error for DecodeError {}

impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: DecodeErrorKind::Wire(error),
        }
    }
}

/// Decode the body-storage and initial-section references from one Pages root.
///
/// The required field-15 base envelope is checked for unique canonical outer
/// framing but its nested payload is deliberately never decoded. Each present
/// reference is strictly preflighted before its singular Buffa lazy `.get()`
/// is forced exactly once.
pub fn decode_document_body(
    source: &[u8],
    options: DecodeOptions,
) -> Result<DocumentBodySnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_document(source, options, &mut budget)?;
    let view: projection::PagesDocumentBodyArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    let body_storage_view = view.body_storage.get()?;
    let initial_section_view = view.initial_section.get()?;
    let projected = DocumentBodySnapshot {
        body_storage: body_storage_view
            .as_ref()
            .map(project_reference)
            .transpose()?,
        initial_section: initial_section_view
            .as_ref()
            .map(project_reference)
            .transpose()?,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

/// Decode the focused Pages document-root facts needed by the migration host.
//
// A complete strict source pass runs before the private Buffa lazy view is
// constructed. The view is used only for the already-selected scalar
// `left_margin`; each singular root reference is forced through the shared
// borrowed `TSP.Reference` view. The TSA super envelope is checked for its
// required nested TSK archive and its language is borrowed directly from the
// source. Repeated page templates remain deferred to
// [`DocumentRootSnapshot::page_templates`].
pub fn decode_document_root<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<DocumentRootSnapshot<'source>, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_root(source, options, &mut budget)?;
    let view: projection::PagesDocumentBodyArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if view.left_margin != strict.left_margin {
        return Err(DecodeError::projection());
    }
    for reference in [
        strict.floating_drawables,
        strict.theme,
        strict.drawables_zorder,
    ]
    .into_iter()
    .flatten()
    {
        cross_check_reference(reference.raw, reference.snapshot, options)?;
    }
    Ok(DocumentRootSnapshot {
        source,
        options,
        floating_drawables: strict
            .floating_drawables
            .map(|reference| reference.snapshot),
        theme: strict.theme.map(|reference| reference.snapshot),
        drawables_zorder: strict.drawables_zorder.map(|reference| reference.snapshot),
        page_template_count: strict.page_template_count,
        left_margin: strict.left_margin,
        document_language: strict.document_language,
    })
}

/// Decode one streamed `TSWP.ObjectAttributeTable.ObjectAttribute` entry.
///
/// The entry is decoded independently so an enclosing repeated section table
/// never enters generated Buffa code or allocates an input-width collection.
pub fn decode_section_boundary(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SectionBoundarySnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_boundary(source, options, &mut budget)?;
    let view: projection::PagesSectionBoundaryEntryLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    if !view.has_character_index() {
        return Err(DecodeError::missing_required(
            "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
        ));
    }
    let section_view = view.section.get()?;
    let projected = SectionBoundarySnapshot {
        character_index: view.character_index,
        section: section_view.as_ref().map(project_reference).transpose()?,
    };
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(strict)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > max_buffa_message_bytes
        || source.len() > options.max_message_bytes
    {
        return Err(buffa::DecodeError::MessageTooLarge.into());
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit());
    }
    Ok(())
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::field_limit(observed, self.max_fields));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_message(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let strict_and_projection = bytes.saturating_mul(2);
        let observed = self.work_bytes.saturating_add(strict_and_projection);
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy)]
struct RootReference<'source> {
    raw: &'source [u8],
    snapshot: ReferenceSnapshot,
}

#[derive(Debug, Clone, Copy)]
struct StrictRoot<'source> {
    floating_drawables: Option<RootReference<'source>>,
    theme: Option<RootReference<'source>>,
    drawables_zorder: Option<RootReference<'source>>,
    page_template_count: usize,
    left_margin: Option<f32>,
    document_language: Option<&'source str>,
}

fn preflight_root<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictRoot<'source>, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend()?;
    let mut floating_drawables = None;
    let mut theme = None;
    let mut drawables_zorder = None;
    let mut page_template_count = 0usize;
    let mut left_margin = None;
    let mut document_language = None;
    let mut saw_super = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            DOCUMENT_FLOATING_DRAWABLES_FIELD => {
                if floating_drawables.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TP.DocumentArchive.floating_drawables",
                    ));
                }
                let raw = field.length_delimited()?;
                let snapshot = preflight_reference(raw, nested_options, budget)?;
                floating_drawables = Some(RootReference { raw, snapshot });
            },
            DOCUMENT_THEME_FIELD => {
                if theme.is_some() {
                    return Err(DecodeError::duplicate_singular("TP.DocumentArchive.theme"));
                }
                let raw = field.length_delimited()?;
                let snapshot = preflight_reference(raw, nested_options, budget)?;
                theme = Some(RootReference { raw, snapshot });
            },
            DOCUMENT_DRAWABLES_ZORDER_FIELD => {
                if drawables_zorder.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TP.DocumentArchive.drawables_zorder",
                    ));
                }
                let raw = field.length_delimited()?;
                let snapshot = preflight_reference(raw, nested_options, budget)?;
                drawables_zorder = Some(RootReference { raw, snapshot });
            },
            DOCUMENT_LEFT_MARGIN_FIELD => {
                if left_margin.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TP.DocumentArchive.left_margin",
                    ));
                }
                left_margin = Some(field.fixed32()?);
            },
            DOCUMENT_PAGE_TEMPLATES_FIELD => {
                page_template_count = page_template_count
                    .checked_add(1)
                    .ok_or_else(|| DecodeError::field_limit(usize::MAX, options.max_fields))?;
                let raw = field.length_delimited()?;
                let _snapshot = preflight_reference(raw, nested_options, budget)?;
            },
            DOCUMENT_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular("TP.DocumentArchive.super"));
                }
                saw_super = true;
                document_language =
                    preflight_tsa_document(field.length_delimited()?, nested_options, budget)?;
            },
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TP.DocumentArchive.super"));
    }
    Ok(StrictRoot {
        floating_drawables,
        theme,
        drawables_zorder,
        page_template_count,
        left_margin,
        document_language,
    })
}

fn preflight_tsa_document<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<&'source str>, DecodeError> {
    budget.charge_message(source.len())?;
    let mut saw_super = false;
    let mut language = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            DOCUMENT_SUPER_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular("TSA.DocumentArchive.super"));
                }
                saw_super = true;
                // TSK.DocumentArchive contains only optional fields. Its
                // payload is deliberately opaque at this boundary.
                let _opaque_tsk_document = field.length_delimited()?;
            },
            DOCUMENT_LANGUAGE_FIELD => {
                if language.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSA.DocumentArchive.document_language",
                    ));
                }
                let bytes = field.length_delimited()?;
                language =
                    Some(std::str::from_utf8(bytes).map_err(|_| buffa::DecodeError::InvalidUtf8)?);
            },
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TSA.DocumentArchive.super"));
    }
    Ok(language)
}

fn preflight_document(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<DocumentBodySnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend()?;
    let mut body_storage = None;
    let mut initial_section = None;
    let mut saw_super = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            DOCUMENT_BODY_STORAGE_FIELD => {
                if body_storage.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TP.DocumentArchive.body_storage",
                    ));
                }
                body_storage = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            DOCUMENT_INITIAL_SECTION_FIELD => {
                if initial_section.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TP.DocumentArchive.section",
                    ));
                }
                initial_section = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            DOCUMENT_SUPER_FIELD => {
                if saw_super {
                    return Err(DecodeError::duplicate_singular("TP.DocumentArchive.super"));
                }
                saw_super = true;
                let _opaque_super = field.length_delimited()?;
            },
            _ => {},
        }
    }
    if !saw_super {
        return Err(DecodeError::missing_required("TP.DocumentArchive.super"));
    }
    Ok(DocumentBodySnapshot {
        body_storage,
        initial_section,
    })
}

fn preflight_boundary(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<SectionBoundarySnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let nested_options = options.descend()?;
    let mut character_index = None;
    let mut section = None;
    let mut saw_section = false;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            BOUNDARY_CHARACTER_INDEX_FIELD => {
                if character_index.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
                    ));
                }
                character_index = Some(u32::try_from(field.varint()?).map_err(|_conversion| {
                    DecodeError::noncanonical("uint32 scalar exceeds u32")
                })?);
            },
            BOUNDARY_SECTION_FIELD => {
                if saw_section {
                    return Err(DecodeError::duplicate_singular(
                        "TSWP.ObjectAttributeTable.ObjectAttribute.object",
                    ));
                }
                saw_section = true;
                section = Some(preflight_reference(
                    field.length_delimited()?,
                    nested_options,
                    budget,
                )?);
            },
            _ => {},
        }
    }
    Ok(SectionBoundarySnapshot {
        character_index: character_index.ok_or_else(|| {
            DecodeError::missing_required(
                "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
            )
        })?,
        section,
    })
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ReferenceSnapshot, DecodeError> {
    budget.charge_message(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                identifier = NonZeroU64::new(field.varint()?)
                    .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))
                    .map(Some)?;
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(decode_int32(require_canonical_int32(field.varint()?)?));
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                if deprecated_is_external.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_is_external",
                    ));
                }
                deprecated_is_external = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    Ok(ReferenceSnapshot {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn cross_check_reference(
    source: &[u8],
    strict: ReferenceSnapshot,
    options: DecodeOptions,
) -> Result<ReferenceSnapshot, DecodeError> {
    let view: projection::ReferenceLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    let projected = project_reference(&view)?;
    if projected != strict {
        return Err(DecodeError::projection());
    }
    Ok(projected)
}

fn project_reference(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<ReferenceSnapshot, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(ReferenceSnapshot {
        identifier: NonZeroU64::new(view.identifier)
            .ok_or_else(|| DecodeError::zero_identifier("TSP.Reference.identifier"))?,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
    })
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(DecodeError::noncanonical("bool scalar is not zero or one")),
    }
}

fn require_canonical_int32(value: u64) -> Result<u64, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    Ok(value)
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Strict preflight proved the u64 is a canonical sign-extended int32."
)]
fn decode_int32(value: u64) -> i32 {
    value as i32
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32([u8; 4]),
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
}

impl<'source> StrictField<'source> {
    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        if self.wire_type != expected {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: expected as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        Ok(())
    }

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Varint)?;
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn fixed32(self) -> Result<f32, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed32)?;
        match self.value {
            StrictValue::Fixed32(bytes) => Ok(f32::from_le_bytes(bytes)),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group => Err(DecodeError::projection()),
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum ParseItem<'source> {
    Field(StrictField<'source>),
    EndGroup(u32),
}

fn next_strict_field<'source>(
    source: &mut &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, options.recursion_limit, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

fn parse_strict_field<'source>(
    source: &mut &'source [u8],
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    if !canonical_key {
        return Err(DecodeError::noncanonical("protobuf field key"));
    }
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let value = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("protobuf varint value"));
            }
            StrictValue::Varint(value)
        },
        buffa::encoding::WireType::Fixed64 => {
            let _bytes = take_exact(source, 8)?;
            StrictValue::Fixed64
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            if !canonical {
                return Err(DecodeError::noncanonical("length-delimited size"));
            }
            let length = usize::try_from(encoded_length)
                .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
            StrictValue::LengthDelimited(take_exact(source, length)?)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(DecodeError::recursion_limit)?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            StrictValue::Group
        },
        buffa::encoding::WireType::EndGroup => {
            return Ok(Some(ParseItem::EndGroup(field_number)));
        },
        buffa::encoding::WireType::Fixed32 => {
            let bytes = take_exact(source, 4)?;
            StrictValue::Fixed32(
                bytes
                    .try_into()
                    .map_err(|_| buffa::DecodeError::UnexpectedEof)?,
            )
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
    })))
}

fn skip_strict_group(
    source: &mut &[u8],
    expected_field_number: u32,
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_varint(source: &mut &[u8]) -> Result<(u64, bool), DecodeError> {
    let original = *source;
    let mut value = 0u64;
    for index in 0..10usize {
        let byte = *original
            .get(index)
            .ok_or(buffa::DecodeError::UnexpectedEof)?;
        if index == 9 && byte > 1 {
            return Err(buffa::DecodeError::VarintTooLong.into());
        }
        value |= u64::from(byte & 0x7f) << (index * 7);
        if byte & 0x80 == 0 {
            let consumed = index + 1;
            *source = &original[consumed..];
            return Ok((value, canonical_varint_len(value) == consumed));
        }
    }
    Err(buffa::DecodeError::VarintTooLong.into())
}

fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (selected, remaining) = source.split_at(length);
    *source = remaining;
    Ok(selected)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::shadow_unrelated,
    reason = "Focused negative tests use explicit panic messages and reuse local fixture roles."
)]
mod tests {
    use buffa::DecodeOptions as BuffaOptions;
    use prost::Message as _;

    use super::*;
    use crate::{tp, tsp, tswp};

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::new(
            source.len().max(1),
            64,
            source.len().saturating_mul(8).max(1),
            8,
        )
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(-7),
            deprecated_is_external: Some(false),
        }
    }

    fn length_delimited(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        let mut tag = u64::from(field) << 3 | 2;
        while tag >= 0x80 {
            output.push((tag as u8 & 0x7f) | 0x80);
            tag >>= 7;
        }
        output.push(tag as u8);
        let mut length = u64::try_from(payload.len()).expect("test payload length");
        while length >= 0x80 {
            output.push((length as u8 & 0x7f) | 0x80);
            length >>= 7;
        }
        output.push(length as u8);
        output.extend_from_slice(payload);
        output
    }

    fn minimal_document(body: Option<&[u8]>, section: Option<&[u8]>) -> Vec<u8> {
        let mut output = Vec::new();
        if let Some(reference) = body {
            output.extend(length_delimited(DOCUMENT_BODY_STORAGE_FIELD, reference));
        }
        if let Some(reference) = section {
            output.extend(length_delimited(DOCUMENT_INITIAL_SECTION_FIELD, reference));
        }
        output.extend([0x7a, 0x00]);
        output
    }

    fn root_super(language: Option<&[u8]>) -> Vec<u8> {
        let mut tsa = length_delimited(DOCUMENT_SUPER_SUPER_FIELD, &[]);
        if let Some(language) = language {
            tsa.extend(length_delimited(DOCUMENT_LANGUAGE_FIELD, language));
        }
        length_delimited(DOCUMENT_SUPER_FIELD, &tsa)
    }

    fn root_reference(field: u32, identifier: u8) -> Vec<u8> {
        length_delimited(field, &[0x08, identifier])
    }

    fn fixed32(field: u32, value: f32) -> Vec<u8> {
        let tag = field << 3 | 5;
        let mut output = Vec::new();
        let mut tag = u64::from(tag);
        while tag >= 0x80 {
            output.push((tag as u8 & 0x7f) | 0x80);
            tag >>= 7;
        }
        output.push(tag as u8);
        output.extend(value.to_le_bytes());
        output
    }

    #[test]
    fn canonical_prost_document_matches_the_strict_projection()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = tp::DocumentArchive {
            body_storage: Some(reference(42)),
            section: Some(reference(99)),
            stylesheet: Some(reference(7)),
            ..tp::DocumentArchive::default()
        }
        .encode_to_vec();
        let snapshot = decode_document_body(&source, options(&source))?;
        assert_eq!(
            snapshot.body_storage().map(ReferenceSnapshot::identifier),
            NonZeroU64::new(42)
        );
        let initial = snapshot.initial_section().expect("initial section");
        assert_eq!(initial.identifier(), NonZeroU64::new(99).expect("non-zero"));
        assert_eq!(initial.deprecated_type(), Some(-7));
        assert_eq!(initial.deprecated_is_external(), Some(false));
        Ok(())
    }

    #[test]
    fn canonical_root_projection_borrows_and_streams_host_facts()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut source = root_super(Some(b"en-US"));
        source.extend(root_reference(DOCUMENT_FLOATING_DRAWABLES_FIELD, 42));
        source.extend(root_reference(DOCUMENT_THEME_FIELD, 43));
        source.extend(root_reference(DOCUMENT_DRAWABLES_ZORDER_FIELD, 44));
        source.extend(fixed32(DOCUMENT_LEFT_MARGIN_FIELD, 12.5));
        source.extend(root_reference(DOCUMENT_PAGE_TEMPLATES_FIELD, 45));
        source.extend(root_reference(DOCUMENT_PAGE_TEMPLATES_FIELD, 46));

        let root = decode_document_root(&source, options(&source))?;
        assert_eq!(root.raw(), source.as_slice());
        assert_eq!(
            root.floating_drawables().map(ReferenceSnapshot::identifier),
            NonZeroU64::new(42)
        );
        assert_eq!(
            root.theme().map(ReferenceSnapshot::identifier),
            NonZeroU64::new(43)
        );
        assert_eq!(
            root.drawables_zorder().map(ReferenceSnapshot::identifier),
            NonZeroU64::new(44)
        );
        assert_eq!(root.left_margin(), Some(12.5));
        assert_eq!(root.document_language(), Some("en-US"));
        assert_eq!(root.page_template_count(), 2);
        let templates = root.page_templates().collect::<Result<Vec<_>, _>>()?;
        assert_eq!(
            templates
                .iter()
                .map(|reference| reference.identifier().get())
                .collect::<Vec<_>>(),
            [45, 46]
        );
        let language = root.document_language().expect("language");
        assert!(language.as_ptr() >= source.as_ptr());
        assert!(language.as_ptr() < source.as_ptr().wrapping_add(source.len()));
        Ok(())
    }

    #[test]
    fn root_projection_rejects_missing_nested_super_and_singular_duplicates() {
        let source: [u8; 0] = [];
        assert_eq!(
            decode_document_root(&source, options(&source))
                .expect_err("missing root super")
                .missing_required_field(),
            Some("TP.DocumentArchive.super")
        );

        let source = [root_super(None), root_super(None)].concat();
        assert_eq!(
            decode_document_root(&source, options(&source))
                .expect_err("duplicate root super")
                .duplicate_singular_field(),
            Some("TP.DocumentArchive.super")
        );

        let source = [0x7a, 0x00];
        let error =
            decode_document_root(&source, options(&source)).expect_err("missing nested TSA super");
        assert_eq!(
            error.missing_required_field(),
            Some("TSA.DocumentArchive.super")
        );

        let mut source = length_delimited(DOCUMENT_SUPER_FIELD, &[]);
        source.extend(root_reference(DOCUMENT_THEME_FIELD, 1));
        let error = decode_document_root(&source, options(&source)).expect_err("missing TSA super");
        assert_eq!(
            error.missing_required_field(),
            Some("TSA.DocumentArchive.super")
        );

        let mut source = root_super(None);
        source.extend(root_reference(DOCUMENT_THEME_FIELD, 1));
        source.extend(root_reference(DOCUMENT_THEME_FIELD, 2));
        let error = decode_document_root(&source, options(&source)).expect_err("duplicate theme");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TP.DocumentArchive.theme")
        );

        let mut source = root_super(None);
        source.extend(fixed32(DOCUMENT_LEFT_MARGIN_FIELD, 1.0));
        source.extend(fixed32(DOCUMENT_LEFT_MARGIN_FIELD, 2.0));
        let error = decode_document_root(&source, options(&source)).expect_err("duplicate margin");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TP.DocumentArchive.left_margin")
        );
    }

    #[test]
    fn root_projection_rejects_bad_reference_wire_values_and_language() {
        let mut source = root_super(None);
        source.extend(length_delimited(DOCUMENT_THEME_FIELD, &[0x08, 0x00]));
        let error = decode_document_root(&source, options(&source)).expect_err("zero theme");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );

        let mut source = root_super(None);
        source.extend(length_delimited(DOCUMENT_THEME_FIELD, &[0x0a, 0x01, 0x01]));
        assert!(decode_document_root(&source, options(&source)).is_err());

        let source = root_super(Some(&[0xff]));
        let error = decode_document_root(&source, options(&source)).expect_err("invalid language");
        assert!(matches!(error.kind, DecodeErrorKind::Wire(_)));

        let mut source = root_super(None);
        source.extend([0x80, 0x02, 0x00]);
        assert!(decode_document_root(&source, options(&source)).is_err());
    }

    #[test]
    fn root_projection_applies_finite_field_and_work_limits() {
        let mut source = root_super(None);
        source.extend(root_reference(DOCUMENT_THEME_FIELD, 1));
        assert!(
            decode_document_root(&source, DecodeOptions::new(source.len(), 0, usize::MAX, 8))
                .is_err()
        );
        assert!(decode_document_root(&source, DecodeOptions::new(source.len(), 64, 1, 8)).is_err());
        assert!(
            decode_document_root(&source, DecodeOptions::new(source.len(), 64, usize::MAX, 0))
                .is_err()
        );
    }

    #[test]
    fn canonical_prost_boundary_matches_the_strict_projection()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = tswp::object_attribute_table::ObjectAttribute {
            character_index: 65_535,
            object: Some(reference(u64::MAX)),
        }
        .encode_to_vec();
        let snapshot = decode_section_boundary(&source, options(&source))?;
        assert_eq!(snapshot.character_index(), 65_535);
        assert_eq!(
            snapshot.section().map(ReferenceSnapshot::identifier),
            NonZeroU64::new(u64::MAX)
        );
        Ok(())
    }

    #[test]
    fn optional_root_and_boundary_references_preserve_absence()
    -> Result<(), Box<dyn std::error::Error>> {
        let document = [0x7a, 0x00];
        assert_eq!(
            decode_document_body(&document, options(&document))?,
            DocumentBodySnapshot {
                body_storage: None,
                initial_section: None,
            }
        );
        let boundary = [0x08, 0x00];
        assert_eq!(
            decode_section_boundary(&boundary, options(&boundary))?,
            SectionBoundarySnapshot {
                character_index: 0,
                section: None,
            }
        );
        Ok(())
    }

    #[test]
    fn opaque_super_payload_is_never_decoded() -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x7a, 0x01, 0xff, 0x22, 0x02, 0x08, 0x2a];
        assert_eq!(
            decode_document_body(&source, options(&source))?
                .body_storage()
                .map(ReferenceSnapshot::identifier),
            NonZeroU64::new(42)
        );
        Ok(())
    }

    #[test]
    fn required_outer_and_nested_fields_are_enforced() {
        let source = [0x22, 0x02, 0x08, 0x2a];
        let error = decode_document_body(&source, options(&source)).expect_err("missing super");
        assert_eq!(
            error.missing_required_field(),
            Some("TP.DocumentArchive.super")
        );

        let source = minimal_document(Some(&[]), None);
        let error = decode_document_body(&source, options(&source)).expect_err("missing id");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let source: [u8; 0] = [];
        let error = decode_section_boundary(&source, options(&source)).expect_err("missing index");
        assert_eq!(
            error.missing_required_field(),
            Some("TSWP.ObjectAttributeTable.ObjectAttribute.character_index")
        );
    }

    #[test]
    fn zero_identifiers_are_rejected_in_every_reference_position() {
        let zero = [0x08, 0x00];
        let document = minimal_document(Some(&zero), None);
        let error = decode_document_body(&document, options(&document)).expect_err("zero body");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );

        let boundary = [0x08, 0x00, 0x12, 0x02, 0x08, 0x00];
        let error =
            decode_section_boundary(&boundary, options(&boundary)).expect_err("zero section");
        assert_eq!(
            error.zero_identifier_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn singular_duplicates_are_rejected_before_buffa_last_one_wins() {
        let mut document = minimal_document(Some(&[0x08, 0x01]), None);
        document.splice(4..4, [0x22, 0x02, 0x08, 0x02]);
        let error =
            decode_document_body(&document, options(&document)).expect_err("duplicate body");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TP.DocumentArchive.body_storage")
        );

        let nested_duplicate = [0x08, 0x01, 0x08, 0x02];
        let document = minimal_document(Some(&nested_duplicate), None);
        let error = decode_document_body(&document, options(&document)).expect_err("duplicate id");
        assert_eq!(
            error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn every_known_singular_scope_rejects_duplicates() {
        let cases: [(Vec<u8>, &str, bool); 6] = [
            (
                [
                    length_delimited(DOCUMENT_INITIAL_SECTION_FIELD, &[0x08, 0x01]),
                    length_delimited(DOCUMENT_INITIAL_SECTION_FIELD, &[0x08, 0x02]),
                    vec![0x7a, 0x00],
                ]
                .concat(),
                "TP.DocumentArchive.section",
                true,
            ),
            (
                [vec![0x7a, 0x00], vec![0x7a, 0x00]].concat(),
                "TP.DocumentArchive.super",
                true,
            ),
            (
                minimal_document(Some(&[0x08, 0x01, 0x10, 0x07, 0x10, 0x08]), None),
                "TSP.Reference.deprecated_type",
                true,
            ),
            (
                minimal_document(Some(&[0x08, 0x01, 0x18, 0x00, 0x18, 0x01]), None),
                "TSP.Reference.deprecated_is_external",
                true,
            ),
            (
                vec![0x08, 0x00, 0x08, 0x01],
                "TSWP.ObjectAttributeTable.ObjectAttribute.character_index",
                false,
            ),
            (
                vec![0x08, 0x00, 0x12, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x02],
                "TSWP.ObjectAttributeTable.ObjectAttribute.object",
                false,
            ),
        ];
        for (source, expected, document) in cases {
            let result = if document {
                decode_document_body(&source, options(&source)).map(|_snapshot| ())
            } else {
                decode_section_boundary(&source, options(&source)).map(|_snapshot| ())
            };
            let Err(error) = result else {
                panic!("{expected} duplicate must fail");
            };
            assert_eq!(error.duplicate_singular_field(), Some(expected));
        }
    }

    #[test]
    fn malformed_body_footnote_corpus_cases_remain_rejected() {
        #[derive(Clone, Copy)]
        enum ExpectedError {
            Missing(&'static str),
            Duplicate(&'static str),
            Noncanonical(&'static str),
            Wire,
        }

        let assert_error = |result: Result<(), DecodeError>, expected: ExpectedError| {
            let error = result.expect_err("malformed corpus case must fail");
            match expected {
                ExpectedError::Missing(field) => {
                    assert_eq!(error.missing_required_field(), Some(field));
                },
                ExpectedError::Duplicate(field) => {
                    assert_eq!(error.duplicate_singular_field(), Some(field));
                },
                ExpectedError::Noncanonical(reason) => {
                    assert_eq!(error.noncanonical_reason(), Some(reason));
                },
                ExpectedError::Wire => {
                    assert!(matches!(&error.kind, DecodeErrorKind::Wire(_)));
                },
            }
        };

        let document_cases: [(&[u8], ExpectedError); 4] = [
            (
                &[0x22, 0x04, 0x08, 0x01, 0x08, 0x02, 0x7a, 0x00],
                ExpectedError::Duplicate("TSP.Reference.identifier"),
            ),
            (
                &[0x22, 0x00, 0x7a, 0x00],
                ExpectedError::Missing("TSP.Reference.identifier"),
            ),
            (
                &[0x22, 0x03, 0x08, 0x80, 0x00, 0x7a, 0x00],
                ExpectedError::Noncanonical("protobuf varint value"),
            ),
            (&[0x22, 0x02, 0x0a, 0x00, 0x7a, 0x00], ExpectedError::Wire),
        ];
        for (source, expected) in document_cases {
            assert_error(
                decode_document_body(source, options(source)).map(|_| ()),
                expected,
            );
        }

        let boundary_cases: [(&[u8], ExpectedError); 4] = [
            (
                &[0x08, 0x01, 0x12, 0x02, 0x08, 0x01, 0x12, 0x02, 0x08, 0x02],
                ExpectedError::Duplicate("TSWP.ObjectAttributeTable.ObjectAttribute.object"),
            ),
            (
                &[0x12, 0x02, 0x08, 0x01],
                ExpectedError::Missing("TSWP.ObjectAttributeTable.ObjectAttribute.character_index"),
            ),
            (
                &[0x08, 0x01, 0x12, 0x00],
                ExpectedError::Missing("TSP.Reference.identifier"),
            ),
            (&[0x08, 0x01, 0x12, 0x02, 0x0a, 0x00], ExpectedError::Wire),
        ];
        for (source, expected) in boundary_cases {
            assert_error(
                decode_section_boundary(source, options(source)).map(|_| ()),
                expected,
            );
        }
    }

    #[test]
    fn malformed_deferred_reference_is_forced() {
        let source = [0x22, 0x01, 0x08, 0x7a, 0x00];
        let direct: projection::PagesDocumentBodyArchiveLazyView<'_> = BuffaOptions::new()
            .with_unknown_field_limit(source.len())
            .decode_lazy_view(&source)
            .expect("outer lazy view");
        assert!(direct.body_storage.get().is_err());
        assert!(decode_document_body(&source, options(&source)).is_err());
    }

    #[test]
    fn wrong_wire_types_are_rejected() {
        let root_wrong_wire = [0x20, 0x2a, 0x7a, 0x00];
        assert!(decode_document_body(&root_wrong_wire, options(&root_wrong_wire)).is_err());

        let nested_wrong_wire = minimal_document(Some(&[0x0a, 0x00]), None);
        assert!(decode_document_body(&nested_wrong_wire, options(&nested_wrong_wire)).is_err());

        let boundary_wrong_wire = [0x0a, 0x00];
        assert!(
            decode_section_boundary(&boundary_wrong_wire, options(&boundary_wrong_wire)).is_err()
        );

        let deprecated_type_wrong_wire = minimal_document(Some(&[0x08, 0x01, 0x12, 0x00]), None);
        assert!(
            decode_document_body(
                &deprecated_type_wrong_wire,
                options(&deprecated_type_wrong_wire)
            )
            .is_err()
        );

        let deprecated_bool_wrong_wire = minimal_document(Some(&[0x08, 0x01, 0x1a, 0x00]), None);
        assert!(
            decode_document_body(
                &deprecated_bool_wrong_wire,
                options(&deprecated_bool_wrong_wire)
            )
            .is_err()
        );
    }

    #[test]
    fn boundary_character_index_rejects_values_above_u32() {
        let source = [0x08, 0x80, 0x80, 0x80, 0x80, 0x10];
        assert_eq!(
            decode_section_boundary(&source, options(&source))
                .expect_err("oversized uint32")
                .noncanonical_reason(),
            Some("uint32 scalar exceeds u32")
        );
    }

    #[test]
    fn noncanonical_keys_lengths_varints_int32_and_bools_are_rejected() {
        let overlong_key = [0xa2, 0x00, 0x02, 0x08, 0x2a, 0x7a, 0x00];
        assert_eq!(
            decode_document_body(&overlong_key, options(&overlong_key))
                .expect_err("overlong key")
                .noncanonical_reason(),
            Some("protobuf field key")
        );

        let overlong_length = [0x22, 0x82, 0x00, 0x08, 0x2a, 0x7a, 0x00];
        assert_eq!(
            decode_document_body(&overlong_length, options(&overlong_length))
                .expect_err("overlong length")
                .noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_identifier = minimal_document(Some(&[0x08, 0xaa, 0x00]), None);
        assert_eq!(
            decode_document_body(&overlong_identifier, options(&overlong_identifier))
                .expect_err("overlong value")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );

        let bad_int32 = minimal_document(
            Some(&[0x08, 0x01, 0x10, 0x80, 0x80, 0x80, 0x80, 0x08]),
            None,
        );
        assert_eq!(
            decode_document_body(&bad_int32, options(&bad_int32))
                .expect_err("bad int32")
                .noncanonical_reason(),
            Some("int32 scalar is not a sign-extended 32-bit value")
        );

        let bad_bool = minimal_document(Some(&[0x08, 0x01, 0x18, 0x02]), None);
        assert_eq!(
            decode_document_body(&bad_bool, options(&bad_bool))
                .expect_err("bad bool")
                .noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );
    }

    #[test]
    fn canonical_unknown_fields_remain_opaque_but_noncanonical_unknowns_fail()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = [0x98, 0x06, 0x01, 0x7a, 0x00];
        assert_eq!(
            decode_document_body(&source, options(&source))?.body_storage(),
            None
        );

        let source = [0x98, 0x06, 0x81, 0x00, 0x7a, 0x00];
        assert_eq!(
            decode_document_body(&source, options(&source))
                .expect_err("noncanonical opaque varint")
                .noncanonical_reason(),
            Some("protobuf varint value")
        );
        Ok(())
    }

    #[test]
    fn exact_bytes_fields_work_and_nesting_limits_are_enforced()
    -> Result<(), Box<dyn std::error::Error>> {
        let root = [0x7a, 0x00];
        decode_document_body(&root, DecodeOptions::new(2, 1, 4, 1))?;
        assert!(decode_document_body(&root, DecodeOptions::new(1, 1, 4, 1)).is_err());
        assert_eq!(
            decode_document_body(&root, DecodeOptions::new(2, 0, 4, 1))
                .expect_err("field cap")
                .field_limit_values(),
            Some((1, 0))
        );
        assert_eq!(
            decode_document_body(&root, DecodeOptions::new(2, 1, 3, 1))
                .expect_err("work cap")
                .work_limit_values(),
            Some((4, 3))
        );

        let nested = minimal_document(Some(&[0x08, 0x01]), None);
        assert!(
            decode_document_body(
                &nested,
                DecodeOptions::new(nested.len(), 3, nested.len() * 8, 0)
            )
            .is_err()
        );
        decode_document_body(
            &nested,
            DecodeOptions::new(nested.len(), 3, nested.len() * 8, 1),
        )?;
        Ok(())
    }

    #[test]
    fn structural_wire_failures_never_panic() {
        let malformed: [&[u8]; 6] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x7a, 0x02, 0x00],
            &[0x0b],
            &[
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
        ];
        for source in malformed {
            assert!(decode_document_body(source, options(source)).is_err());
            assert!(decode_section_boundary(source, options(source)).is_err());
        }
    }
}
