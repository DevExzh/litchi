//! Private-type Buffa projection for `KN.ShowArchive`.
//!
//! A strict, allocation-free wire preflight completes before Buffa sees the
//! payload. The generated lazy view then projects only direct references,
//! presentation size, and scalar settings. The repeated slide tree is routed
//! by hand, so untrusted slide count never creates a generated repeated-field
//! vector. Caller-owned bytes remain authoritative for preservation.

use std::fmt;

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_keynote_show_generated::LitchiIwaProjection as projection;

const SHOW_UI_STATE_FIELD: u32 = 1;
const SHOW_THEME_FIELD: u32 = 2;
const SHOW_SLIDE_TREE_FIELD: u32 = 3;
const SHOW_SIZE_FIELD: u32 = 4;
const SHOW_STYLESHEET_FIELD: u32 = 5;
const SHOW_SLIDE_NUMBERS_VISIBLE_FIELD: u32 = 6;
const SHOW_RECORDING_FIELD: u32 = 7;
const SHOW_LOOP_PRESENTATION_FIELD: u32 = 8;
const SHOW_MODE_FIELD: u32 = 9;
const SHOW_AUTOPLAY_TRANSITION_DELAY_FIELD: u32 = 10;
const SHOW_AUTOPLAY_BUILD_DELAY_FIELD: u32 = 11;
const SHOW_IDLE_TIMER_ACTIVE_FIELD: u32 = 15;
const SHOW_IDLE_TIMER_DELAY_FIELD: u32 = 16;
const SHOW_SOUNDTRACK_FIELD: u32 = 17;
const SHOW_AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD: u32 = 18;
const SHOW_SLIDE_LIST_FIELD: u32 = 19;

const SLIDE_TREE_ROOT_FIELD: u32 = 1;
const SLIDE_TREE_SLIDES_FIELD: u32 = 2;

const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;

const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;

const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;
const MAX_RECURSION_LIMIT: u32 = 64;

/// Finite resource profile for one Keynote show projection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_slide_references: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
}

impl DecodeOptions {
    /// Build an explicit finite profile.
    ///
    /// The compatibility constructor derives conservative field and aggregate
    /// scan-work ceilings from the byte ceiling. Callers that already have
    /// independent wire limits can replace those two ceilings with the
    /// `with_max_*` builders.
    #[must_use]
    pub const fn new(
        max_message_bytes: usize,
        max_slide_references: usize,
        recursion_limit: u32,
    ) -> Self {
        Self {
            max_message_bytes,
            max_slide_references,
            max_fields: max_message_bytes.saturating_mul(4),
            max_work_bytes: max_message_bytes.saturating_mul(8),
            recursion_limit,
        }
    }

    /// Replace the aggregate handwritten field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, max_fields: usize) -> Self {
        self.max_fields = max_fields;
        self
    }

    /// Replace the aggregate strict-plus-Buffa scan-work ceiling in bytes.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, max_work_bytes: usize) -> Self {
        self.max_work_bytes = max_work_bytes;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self, budget: &Budget) -> Result<Self, DecodeError> {
        if self.recursion_limit <= 1 {
            return Err(budget.nesting_limit());
        }
        Ok(Self {
            recursion_limit: self.recursion_limit - 1,
            ..self
        })
    }
}

/// Failure from the bounded Keynote show projection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

/// A content-free exact byte or nesting resource failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum WireResourceLimit {
    /// Complete input bytes exceeded the configured or Buffa hard ceiling.
    Bytes {
        /// Exact input or configured value that exceeded the ceiling.
        observed: usize,
        /// Exact applied ceiling.
        maximum: usize,
    },
    /// Configured or traversed nesting exceeded its exact ceiling.
    Nesting {
        /// Exact configured value or first rejected depth.
        observed: u32,
        /// Exact applied ceiling.
        maximum: u32,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Resource(WireResourceLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    SlideReferenceLimit { observed: usize, maximum: usize },
    FieldLimit { observed: usize, maximum: usize },
    WorkLimit { observed: usize, maximum: usize },
    Allocation { amount: usize },
    Projection,
}

impl DecodeError {
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

    const fn slide_reference_limit(observed: usize, maximum: usize) -> Self {
        Self {
            kind: DecodeErrorKind::SlideReferenceLimit { observed, maximum },
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

    const fn allocation(amount: usize) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation { amount },
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    /// Required schema field absent from the source, when applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::SlideReferenceLimit { .. }
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Allocation { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Singular schema field repeated in the source, when applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::SlideReferenceLimit { .. }
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Allocation { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Stable reason for a non-canonical known wire representation.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::SlideReferenceLimit { .. }
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Allocation { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Observed and maximum slide-reference counts for a resource failure.
    #[must_use]
    pub const fn slide_reference_limit_values(&self) -> Option<(usize, usize)> {
        match self.kind {
            DecodeErrorKind::SlideReferenceLimit { observed, maximum } => Some((observed, maximum)),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Allocation { .. }
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Exact observed and configured field counts for a field-limit failure.
    #[must_use]
    pub const fn field_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::FieldLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Exact observed and configured work bytes for a work-limit failure.
    #[must_use]
    pub const fn work_limit_values(&self) -> Option<(usize, usize)> {
        let DecodeErrorKind::WorkLimit { observed, maximum } = self.kind else {
            return None;
        };
        Some((observed, maximum))
    }

    /// Exact byte or nesting resource failure, when applicable.
    #[must_use]
    pub const fn wire_resource_limit(&self) -> Option<WireResourceLimit> {
        let DecodeErrorKind::Resource(limit) = self.kind else {
            return None;
        };
        Some(limit)
    }

    /// Requested retained-reference capacity for an allocation failure.
    #[must_use]
    pub const fn allocation_amount(&self) -> Option<usize> {
        match self.kind {
            DecodeErrorKind::Allocation { amount } => Some(amount),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Resource(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::SlideReferenceLimit { .. }
            | DecodeErrorKind::FieldLimit { .. }
            | DecodeErrorKind::WorkLimit { .. }
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Resource(WireResourceLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "Keynote show projection byte limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::Resource(WireResourceLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "Keynote show projection nesting limit exceeded: observed {observed}, maximum {maximum}"
            ),
            DecodeErrorKind::MissingRequired(field) => {
                write!(formatter, "missing required field {field}")
            },
            DecodeErrorKind::DuplicateSingular(field) => {
                write!(formatter, "duplicate singular field {field}")
            },
            DecodeErrorKind::NonCanonical(reason) => {
                write!(formatter, "non-canonical protobuf representation: {reason}")
            },
            DecodeErrorKind::SlideReferenceLimit { observed, maximum } => write!(
                formatter,
                "Keynote show has {observed} slide references; maximum is {maximum}"
            ),
            DecodeErrorKind::FieldLimit { observed, maximum } => write!(
                formatter,
                "Keynote show projection visited {observed} fields; maximum is {maximum}"
            ),
            DecodeErrorKind::WorkLimit { observed, maximum } => write!(
                formatter,
                "Keynote show projection requires {observed} work bytes; maximum is {maximum}"
            ),
            DecodeErrorKind::Allocation { amount } => write!(
                formatter,
                "cannot allocate Keynote slide-reference capacity for {amount} entries"
            ),
            DecodeErrorKind::Projection => {
                formatter.write_str("Keynote show strict preflight disagrees with Buffa projection")
            },
        }
    }
}

impl std::error::Error for DecodeError {}

#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "All present and future non-resource Buffa failures retain their exact wire error."
)]
impl From<buffa::DecodeError> for DecodeError {
    fn from(error: buffa::DecodeError) -> Self {
        Self {
            kind: match error {
                buffa::DecodeError::MessageTooLarge
                | buffa::DecodeError::RecursionLimitExceeded => DecodeErrorKind::Projection,
                other => DecodeErrorKind::Wire(other),
            },
        }
    }
}

/// Raw finite-width presentation size from `TSP.Size`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RawSize {
    width: f32,
    height: f32,
}

impl RawSize {
    /// Native width scalar.
    #[must_use]
    pub const fn width(self) -> f32 {
        self.width
    }

    /// Native height scalar.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.height
    }
}

/// Optional scalar settings retained from `KN.ShowArchive`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RawSettings {
    slide_numbers_visible: Option<bool>,
    loop_presentation: Option<bool>,
    mode: Option<i32>,
    autoplay_transition_delay: Option<f64>,
    autoplay_build_delay: Option<f64>,
    idle_timer_active: Option<bool>,
    idle_timer_delay: Option<f64>,
    automatically_plays_upon_open: Option<bool>,
}

impl RawSettings {
    /// Optional native slide-number visibility flag.
    #[must_use]
    pub const fn slide_numbers_visible(self) -> Option<bool> {
        self.slide_numbers_visible
    }

    /// Optional native presentation-loop flag.
    #[must_use]
    pub const fn loop_presentation(self) -> Option<bool> {
        self.loop_presentation
    }

    /// Optional raw show-mode integer, including unknown future values.
    #[must_use]
    pub const fn mode_raw(self) -> Option<i32> {
        self.mode
    }

    /// Optional native autoplay transition delay.
    #[must_use]
    pub const fn autoplay_transition_delay(self) -> Option<f64> {
        self.autoplay_transition_delay
    }

    /// Optional native autoplay build delay.
    #[must_use]
    pub const fn autoplay_build_delay(self) -> Option<f64> {
        self.autoplay_build_delay
    }

    /// Optional native idle-timer activation flag.
    #[must_use]
    pub const fn idle_timer_active(self) -> Option<bool> {
        self.idle_timer_active
    }

    /// Optional native idle-timer delay.
    #[must_use]
    pub const fn idle_timer_delay(self) -> Option<f64> {
        self.idle_timer_delay
    }

    /// Optional native play-on-open flag.
    #[must_use]
    pub const fn automatically_plays_upon_open(self) -> Option<bool> {
        self.automatically_plays_upon_open
    }
}

/// Owned, generated-type-free projection of Keynote presentation settings.
///
/// Unlike [`ShowSnapshot`], this value does not retain the presentation's
/// slide-node identifiers. The complete known show envelope and slide tree
/// are still checked by the strict bounded preflight before Buffa projects the
/// size and scalar settings.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SettingsSnapshot {
    size: RawSize,
    raw_settings: RawSettings,
}

impl SettingsSnapshot {
    /// Raw native presentation size.
    #[must_use]
    pub const fn size(self) -> RawSize {
        self.size
    }

    /// Optional scalar show settings.
    #[must_use]
    pub const fn raw_settings(self) -> RawSettings {
        self.raw_settings
    }
}

/// Owned, generated-type-free Keynote show projection.
#[derive(Debug, Clone, PartialEq)]
pub struct ShowSnapshot {
    slide_node_identifiers: Box<[u64]>,
    size: RawSize,
    raw_settings: RawSettings,
    has_deprecated_root_slide_node: bool,
    has_slide_list: bool,
}

impl ShowSnapshot {
    /// Ordered slide-node identifiers from the native slide tree.
    #[must_use]
    pub fn slide_node_identifiers(&self) -> &[u64] {
        &self.slide_node_identifiers
    }

    /// Raw native presentation size.
    #[must_use]
    pub const fn size(&self) -> RawSize {
        self.size
    }

    /// Optional scalar show settings.
    #[must_use]
    pub const fn raw_settings(&self) -> RawSettings {
        self.raw_settings
    }

    /// Whether the deprecated root-slide-node topology field is present.
    #[must_use]
    pub const fn has_deprecated_root_slide_node(&self) -> bool {
        self.has_deprecated_root_slide_node
    }

    /// Whether the secondary grouped slide-list topology field is present.
    #[must_use]
    pub const fn has_slide_list(&self) -> bool {
        self.has_slide_list
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RawReference {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

#[derive(Debug)]
struct ShowPreflight<'source> {
    slide_tree: &'source [u8],
    slide_count: usize,
    projected_nested_bytes: usize,
    size: RawSize,
    raw_settings: RawSettings,
    references: ShowReferences,
    has_deprecated_root_slide_node: bool,
    has_slide_list: bool,
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_nesting: u32,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_nesting: options.recursion_limit,
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

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes);
        if observed > self.max_work_bytes {
            return Err(DecodeError::work_limit(observed, self.max_work_bytes));
        }
        self.work_bytes = observed;
        Ok(())
    }

    const fn nesting_limit(&self) -> DecodeError {
        DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: self.max_nesting.saturating_add(1),
                maximum: self.max_nesting,
            }),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ShowReferences {
    ui_state: Option<RawReference>,
    theme: RawReference,
    stylesheet: RawReference,
    recording: Option<RawReference>,
    soundtrack: Option<RawReference>,
    slide_list: Option<RawReference>,
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64(u64),
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32(u32),
}

#[derive(Clone, Copy, Debug)]
struct StrictField<'source> {
    number: u32,
    wire_type: buffa::encoding::WireType,
    value: StrictValue<'source>,
    canonical_key: bool,
    canonical_value: bool,
}

impl<'source> StrictField<'source> {
    fn require_canonical_key(self) -> Result<(), DecodeError> {
        if !self.canonical_key {
            return Err(DecodeError::noncanonical("protobuf field key"));
        }
        Ok(())
    }

    fn require_wire_type(self, expected: buffa::encoding::WireType) -> Result<(), DecodeError> {
        self.require_canonical_key()?;
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
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::LengthDelimited)?;
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
        }
    }

    fn fixed32_bits(self) -> Result<u32, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed32)?;
        match self.value {
            StrictValue::Fixed32(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group => Err(DecodeError::projection()),
        }
    }

    fn fixed64_bits(self) -> Result<u64, DecodeError> {
        self.require_wire_type(buffa::encoding::WireType::Fixed64)?;
        match self.value {
            StrictValue::Fixed64(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32(_) => Err(DecodeError::projection()),
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
    recursion_limit: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, budget)? {
        Some(ParseItem::Field(field)) => Ok(Some(field)),
        Some(ParseItem::EndGroup(number)) => {
            Err(buffa::DecodeError::InvalidEndGroup(number).into())
        },
        None => Ok(None),
    }
}

/// Decode one bounded Keynote show payload.
///
/// The strict handwritten pass validates the complete known envelope and
/// slide tree before Buffa records even one lazy fragment. Slide references
/// are then retained after one exact fallible reservation and projected in
/// source order. Generated Buffa types never cross this API boundary.
pub fn decode_show(source: &[u8], options: DecodeOptions) -> Result<ShowSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let preflight = preflight_show(source, options, &mut budget)?;
    let mut slide_node_identifiers = Vec::new();
    slide_node_identifiers
        .try_reserve_exact(preflight.slide_count)
        .map_err(|_error| DecodeError::allocation(preflight.slide_count))?;

    let settings = project_settings(source, options, &preflight, &mut budget)?;

    project_slide_tree(
        preflight.slide_tree,
        options.descend(&budget)?,
        preflight.slide_count,
        &mut slide_node_identifiers,
        &mut budget,
    )?;
    Ok(ShowSnapshot {
        slide_node_identifiers: slide_node_identifiers.into_boxed_slice(),
        size: settings.size,
        raw_settings: settings.raw_settings,
        has_deprecated_root_slide_node: preflight.has_deprecated_root_slide_node,
        has_slide_list: preflight.has_slide_list,
    })
}

/// Decode only the bounded Keynote presentation settings projection.
///
/// The strict handwritten pass validates the complete known show envelope and
/// slide tree, including the caller's slide-reference ceiling. Buffa then
/// projects the presentation size and scalar settings without allocating or
/// retaining the slide-node identifier list.
pub fn decode_settings(
    source: &[u8],
    options: DecodeOptions,
) -> Result<SettingsSnapshot, DecodeError> {
    validate_decode_input(source, options)?;
    let mut budget = Budget::new(options);
    let preflight = preflight_show(source, options, &mut budget)?;
    project_settings(source, options, &preflight, &mut budget)
}

fn validate_decode_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let max_buffa_message_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| DecodeError::projection())?;
    if options.max_message_bytes > max_buffa_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: options.max_message_bytes,
                maximum: max_buffa_message_bytes,
            }),
        });
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: options.max_message_bytes,
            }),
        });
    }
    if options.recursion_limit == 0 || options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError {
            kind: DecodeErrorKind::Resource(WireResourceLimit::Nesting {
                observed: options.recursion_limit,
                maximum: MAX_RECURSION_LIMIT,
            }),
        });
    }
    Ok(())
}

fn project_settings(
    source: &[u8],
    options: DecodeOptions,
    preflight: &ShowPreflight<'_>,
    budget: &mut Budget,
) -> Result<SettingsSnapshot, DecodeError> {
    budget.charge_work(source.len())?;
    let view: projection::KeynoteShowArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    budget.charge_work(preflight.projected_nested_bytes)?;
    let (projected_size, projected_references) = force_show_projection(&view)?;
    let projected_settings = RawSettings {
        slide_numbers_visible: view.slide_numbers_visible,
        loop_presentation: view.loop_presentation,
        mode: view.mode,
        autoplay_transition_delay: view.autoplay_transition_delay,
        autoplay_build_delay: view.autoplay_build_delay,
        idle_timer_active: view.idle_timer_active,
        idle_timer_delay: view.idle_timer_delay,
        automatically_plays_upon_open: view.automatically_plays_upon_open,
    };
    if projected_size.width.to_bits() != preflight.size.width.to_bits()
        || projected_size.height.to_bits() != preflight.size.height.to_bits()
        || projected_references != preflight.references
        || !same_raw_settings(projected_settings, preflight.raw_settings)
    {
        return Err(DecodeError::projection());
    }
    Ok(SettingsSnapshot {
        size: projected_size,
        raw_settings: projected_settings,
    })
}

fn preflight_show<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<ShowPreflight<'source>, DecodeError> {
    budget.charge_work(source.len())?;
    let nested_options = options.descend(budget)?;
    let mut seen = 0u32;
    let mut slide_tree = None;
    let mut size = None;
    let mut ui_state = None;
    let mut theme = None;
    let mut stylesheet = None;
    let mut recording = None;
    let mut soundtrack = None;
    let mut slide_list = None;
    let mut raw_settings = RawSettings {
        slide_numbers_visible: None,
        loop_presentation: None,
        mode: None,
        autoplay_transition_delay: None,
        autoplay_build_delay: None,
        idle_timer_active: None,
        idle_timer_delay: None,
        automatically_plays_upon_open: None,
    };
    let mut has_slide_list = false;
    let mut projected_nested_bytes = 0usize;

    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            SHOW_UI_STATE_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.uiState")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                ui_state = Some(preflight_reference(payload, nested_options, budget)?);
            },
            SHOW_THEME_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.theme")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                theme = Some(preflight_reference(payload, nested_options, budget)?);
            },
            SHOW_SLIDE_TREE_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.slideTree")?;
                slide_tree = Some(field.length_delimited()?);
            },
            SHOW_SIZE_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.size")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                size = Some(preflight_size(payload, nested_options, budget)?);
            },
            SHOW_STYLESHEET_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.stylesheet")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                stylesheet = Some(preflight_reference(payload, nested_options, budget)?);
            },
            SHOW_SLIDE_NUMBERS_VISIBLE_FIELD => {
                mark_singular(
                    &mut seen,
                    field.number,
                    "KN.ShowArchive.slideNumbersVisible",
                )?;
                raw_settings.slide_numbers_visible = Some(require_canonical_bool(field.varint()?)?);
            },
            SHOW_RECORDING_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.recording")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                recording = Some(preflight_reference(payload, nested_options, budget)?);
            },
            SHOW_LOOP_PRESENTATION_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.loop_presentation")?;
                raw_settings.loop_presentation = Some(require_canonical_bool(field.varint()?)?);
            },
            SHOW_MODE_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.mode")?;
                raw_settings.mode = Some(require_canonical_int32(field.varint()?)?);
            },
            SHOW_AUTOPLAY_TRANSITION_DELAY_FIELD => {
                mark_singular(
                    &mut seen,
                    field.number,
                    "KN.ShowArchive.autoplay_transition_delay",
                )?;
                raw_settings.autoplay_transition_delay =
                    Some(f64::from_bits(field.fixed64_bits()?));
            },
            SHOW_AUTOPLAY_BUILD_DELAY_FIELD => {
                mark_singular(
                    &mut seen,
                    field.number,
                    "KN.ShowArchive.autoplay_build_delay",
                )?;
                raw_settings.autoplay_build_delay = Some(f64::from_bits(field.fixed64_bits()?));
            },
            SHOW_IDLE_TIMER_ACTIVE_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.idle_timer_active")?;
                raw_settings.idle_timer_active = Some(require_canonical_bool(field.varint()?)?);
            },
            SHOW_IDLE_TIMER_DELAY_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.idle_timer_delay")?;
                raw_settings.idle_timer_delay = Some(f64::from_bits(field.fixed64_bits()?));
            },
            SHOW_SOUNDTRACK_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.soundtrack")?;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                soundtrack = Some(preflight_reference(payload, nested_options, budget)?);
            },
            SHOW_AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD => {
                mark_singular(
                    &mut seen,
                    field.number,
                    "KN.ShowArchive.automatically_plays_upon_open",
                )?;
                raw_settings.automatically_plays_upon_open =
                    Some(require_canonical_bool(field.varint()?)?);
            },
            SHOW_SLIDE_LIST_FIELD => {
                mark_singular(&mut seen, field.number, "KN.ShowArchive.slideList")?;
                has_slide_list = true;
                let payload = field.length_delimited()?;
                projected_nested_bytes = add_nested_bytes(projected_nested_bytes, payload.len())?;
                slide_list = Some(preflight_reference(payload, nested_options, budget)?);
            },
            _ => {},
        }
    }

    require_seen(seen, SHOW_THEME_FIELD, "KN.ShowArchive.theme")?;
    require_seen(seen, SHOW_SLIDE_TREE_FIELD, "KN.ShowArchive.slideTree")?;
    require_seen(seen, SHOW_SIZE_FIELD, "KN.ShowArchive.size")?;
    require_seen(seen, SHOW_STYLESHEET_FIELD, "KN.ShowArchive.stylesheet")?;
    let selected_slide_tree =
        slide_tree.ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.slideTree"))?;
    let selected_size = size.ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.size"))?;
    let (slide_count, has_deprecated_root_slide_node) =
        preflight_slide_tree(selected_slide_tree, nested_options, budget)?;
    Ok(ShowPreflight {
        slide_tree: selected_slide_tree,
        slide_count,
        projected_nested_bytes,
        size: selected_size,
        raw_settings,
        references: ShowReferences {
            ui_state,
            theme: theme.ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.theme"))?,
            stylesheet: stylesheet
                .ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.stylesheet"))?,
            recording,
            soundtrack,
            slide_list,
        },
        has_deprecated_root_slide_node,
        has_slide_list,
    })
}

fn preflight_slide_tree(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<(usize, bool), DecodeError> {
    budget.charge_work(source.len())?;
    let reference_options = options.descend(budget)?;
    let mut root_seen = false;
    let mut slide_count = 0usize;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            SLIDE_TREE_ROOT_FIELD => {
                if root_seen {
                    return Err(DecodeError::duplicate_singular(
                        "KN.SlideTreeArchive.rootSlideNode",
                    ));
                }
                root_seen = true;
                preflight_reference(field.length_delimited()?, reference_options, budget)?;
            },
            SLIDE_TREE_SLIDES_FIELD => {
                preflight_reference(field.length_delimited()?, reference_options, budget)?;
                slide_count = slide_count.checked_add(1).ok_or_else(|| {
                    DecodeError::slide_reference_limit(usize::MAX, options.max_slide_references)
                })?;
                if slide_count > options.max_slide_references {
                    return Err(DecodeError::slide_reference_limit(
                        slide_count,
                        options.max_slide_references,
                    ));
                }
            },
            _ => {},
        }
    }
    Ok((slide_count, root_seen))
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RawReference, DecodeError> {
    budget.charge_work(source.len())?;
    let mut seen = 0u32;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                mark_singular(&mut seen, field.number, "TSP.Reference.identifier")?;
                identifier = Some(field.varint()?);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                mark_singular(&mut seen, field.number, "TSP.Reference.deprecated_type")?;
                deprecated_type = Some(require_canonical_int32(field.varint()?)?);
            },
            REFERENCE_DEPRECATED_EXTERNAL_FIELD => {
                mark_singular(
                    &mut seen,
                    field.number,
                    "TSP.Reference.deprecated_is_external",
                )?;
                deprecated_is_external = Some(require_canonical_bool(field.varint()?)?);
            },
            _ => {},
        }
    }
    require_seen(seen, REFERENCE_IDENTIFIER_FIELD, "TSP.Reference.identifier")?;
    Ok(RawReference {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn preflight_size(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RawSize, DecodeError> {
    budget.charge_work(source.len())?;
    let mut seen = 0u32;
    let mut width = None;
    let mut height = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        match field.number {
            SIZE_WIDTH_FIELD => {
                mark_singular(&mut seen, field.number, "TSP.Size.width")?;
                width = Some(f32::from_bits(field.fixed32_bits()?));
            },
            SIZE_HEIGHT_FIELD => {
                mark_singular(&mut seen, field.number, "TSP.Size.height")?;
                height = Some(f32::from_bits(field.fixed32_bits()?));
            },
            _ => {},
        }
    }
    require_seen(seen, SIZE_WIDTH_FIELD, "TSP.Size.width")?;
    require_seen(seen, SIZE_HEIGHT_FIELD, "TSP.Size.height")?;
    Ok(RawSize {
        width: width.ok_or_else(|| DecodeError::missing_required("TSP.Size.width"))?,
        height: height.ok_or_else(|| DecodeError::missing_required("TSP.Size.height"))?,
    })
}

fn force_show_projection(
    view: &projection::KeynoteShowArchiveLazyView<'_>,
) -> Result<(RawSize, ShowReferences), DecodeError> {
    let ui_state = view
        .ui_state
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    let theme_view = view
        .theme
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.theme"))?;
    let theme = force_reference_projection(&theme_view)?;
    let size = view
        .size
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.size"))?;
    let raw_size = force_size_projection(&size)?;
    let stylesheet_view = view
        .stylesheet
        .get()?
        .ok_or_else(|| DecodeError::missing_required("KN.ShowArchive.stylesheet"))?;
    let stylesheet = force_reference_projection(&stylesheet_view)?;
    let recording = view
        .recording
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    let soundtrack = view
        .soundtrack
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    let slide_list = view
        .slide_list
        .get()?
        .as_ref()
        .map(force_reference_projection)
        .transpose()?;
    Ok((
        raw_size,
        ShowReferences {
            ui_state,
            theme,
            stylesheet,
            recording,
            soundtrack,
            slide_list,
        },
    ))
}

fn same_raw_settings(left: RawSettings, right: RawSettings) -> bool {
    left.slide_numbers_visible == right.slide_numbers_visible
        && left.loop_presentation == right.loop_presentation
        && left.mode == right.mode
        && left.autoplay_transition_delay.map(f64::to_bits)
            == right.autoplay_transition_delay.map(f64::to_bits)
        && left.autoplay_build_delay.map(f64::to_bits)
            == right.autoplay_build_delay.map(f64::to_bits)
        && left.idle_timer_active == right.idle_timer_active
        && left.idle_timer_delay.map(f64::to_bits) == right.idle_timer_delay.map(f64::to_bits)
        && left.automatically_plays_upon_open == right.automatically_plays_upon_open
}

fn add_nested_bytes(total: usize, amount: usize) -> Result<usize, DecodeError> {
    total
        .checked_add(amount)
        .ok_or_else(DecodeError::projection)
}

fn force_reference_projection(
    view: &projection::ReferenceLazyView<'_>,
) -> Result<RawReference, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(RawReference {
        identifier: view.identifier,
        deprecated_type: view.deprecated_type,
        deprecated_is_external: view.deprecated_is_external,
    })
}

fn force_size_projection(view: &projection::SizeLazyView<'_>) -> Result<RawSize, DecodeError> {
    if !view.has_width() {
        return Err(DecodeError::missing_required("TSP.Size.width"));
    }
    if !view.has_height() {
        return Err(DecodeError::missing_required("TSP.Size.height"));
    }
    Ok(RawSize {
        width: view.width,
        height: view.height,
    })
}

fn project_slide_tree(
    source: &[u8],
    options: DecodeOptions,
    expected_slide_count: usize,
    output: &mut Vec<u64>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    budget.charge_work(source.len())?;
    let reference_options = options.descend(budget)?;
    let mut remaining = source;
    while let Some(field) = next_strict_field(&mut remaining, options.recursion_limit, budget)? {
        if field.number != SLIDE_TREE_ROOT_FIELD && field.number != SLIDE_TREE_SLIDES_FIELD {
            continue;
        }
        let payload = field.length_delimited()?;
        let strict = preflight_reference(payload, reference_options, budget)?;
        budget.charge_work(payload.len())?;
        let projected_view: projection::ReferenceLazyView<'_> =
            reference_options.buffa().decode_lazy_view(payload)?;
        let projected_reference = force_reference_projection(&projected_view)?;
        if projected_reference != strict {
            return Err(DecodeError::projection());
        }
        if field.number == SLIDE_TREE_SLIDES_FIELD {
            output.push(projected_reference.identifier);
        }
    }
    if output.len() != expected_slide_count {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn mark_singular(seen: &mut u32, field_number: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u32
        .checked_shl(field_number)
        .ok_or_else(|| DecodeError::noncanonical("known field number exceeds presence mask"))?;
    if *seen & bit != 0 {
        return Err(DecodeError::duplicate_singular(name));
    }
    *seen |= bit;
    Ok(())
}

fn require_seen(seen: u32, field_number: u32, name: &'static str) -> Result<(), DecodeError> {
    let bit = 1u32
        .checked_shl(field_number)
        .ok_or_else(|| DecodeError::noncanonical("known field number exceeds presence mask"))?;
    if seen & bit == 0 {
        return Err(DecodeError::missing_required(name));
    }
    Ok(())
}

fn require_canonical_bool(value: u64) -> Result<bool, DecodeError> {
    if value > 1 {
        return Err(DecodeError::noncanonical("bool scalar is not zero or one"));
    }
    Ok(value == 1)
}

fn require_canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value > 0x7fff_ffff && value < MIN_SIGN_EXTENDED_INT32 {
        return Err(DecodeError::noncanonical(
            "int32 scalar is not a sign-extended 32-bit value",
        ));
    }
    Ok(decode_int32(value))
}

#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_possible_wrap,
    reason = "Strict preflight proved the u64 is a canonical sign-extended int32."
)]
fn decode_int32(value: u64) -> i32 {
    value as i32
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
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_error| buffa::DecodeError::InvalidFieldNumber)?;
    let field_number = raw_tag >> 3;
    if field_number == 0 || field_number > buffa::encoding::MAX_FIELD_NUMBER {
        return Err(buffa::DecodeError::InvalidFieldNumber.into());
    }
    let raw_wire_type = raw_tag & 7;
    let wire_type = buffa::encoding::WireType::from_u32(raw_wire_type)?;
    let (value, canonical_value) = match wire_type {
        buffa::encoding::WireType::Varint => {
            let (value, canonical) = take_varint(source)?;
            (StrictValue::Varint(value), canonical)
        },
        buffa::encoding::WireType::Fixed64 => {
            let bytes = take_exact(source, 8)?;
            let bits = u64::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_error| buffa::DecodeError::UnexpectedEof)?,
            );
            (StrictValue::Fixed64(bits), true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_error| buffa::DecodeError::MessageTooLarge)?;
            (
                StrictValue::LengthDelimited(take_exact(source, length)?),
                canonical,
            )
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit
                .checked_sub(1)
                .ok_or_else(|| budget.nesting_limit())?;
            skip_strict_group(source, field_number, child_limit, budget)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            let bytes = take_exact(source, 4)?;
            let bits = u32::from_le_bytes(
                bytes
                    .try_into()
                    .map_err(|_error| buffa::DecodeError::UnexpectedEof)?,
            );
            (StrictValue::Fixed32(bits), true)
        },
        _ => return Err(buffa::DecodeError::InvalidWireType(raw_wire_type).into()),
    };
    Ok(Some(ParseItem::Field(StrictField {
        number: field_number,
        wire_type,
        value,
        canonical_key,
        canonical_value,
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
            Some(ParseItem::EndGroup(field_number)) if field_number == expected_field_number => {
                return Ok(());
            },
            Some(ParseItem::EndGroup(field_number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(field_number).into());
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
mod tests {
    use prost::Message as _;

    use super::*;
    use crate::{kn, tsp};

    const NATIVE_SHOW: [u8; 80] = [
        0x12, 0x05, 0x08, 0xce, 0xe7, 0xa1, 0x01, 0x1a, 0x07, 0x12, 0x05, 0x08, 0xf5, 0xef, 0xa1,
        0x01, 0x22, 0x0a, 0x0d, 0x00, 0x00, 0xf0, 0x44, 0x15, 0x00, 0x00, 0x87, 0x44, 0x2a, 0x05,
        0x08, 0xe4, 0xf1, 0xa1, 0x01, 0x40, 0x00, 0x48, 0x00, 0x51, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x00, 0x14, 0x40, 0x59, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x40, 0x78, 0x00, 0x81,
        0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x20, 0x8c, 0x40, 0x8a, 0x01, 0x05, 0x08, 0xd5, 0xe7,
        0xa1, 0x01, 0x90, 0x01, 0x00,
    ];

    fn options(source: &[u8], max_slide_references: usize) -> DecodeOptions {
        DecodeOptions::new(source.len().max(1), max_slide_references, 8)
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(7),
            deprecated_is_external: Some(false),
        }
    }

    #[allow(
        deprecated,
        reason = "The regression intentionally exercises deprecated rootSlideNode presence."
    )]
    fn show(slides: &[u64]) -> kn::ShowArchive {
        kn::ShowArchive {
            ui_state: Some(reference(10)),
            theme: reference(11),
            slide_tree: kn::SlideTreeArchive {
                root_slide_node: Some(reference(12)),
                slides: slides.iter().copied().map(reference).collect(),
            },
            size: tsp::Size {
                width: 1920.0,
                height: 1080.0,
            },
            stylesheet: reference(13),
            slide_numbers_visible: Some(true),
            recording: Some(reference(14)),
            loop_presentation: Some(false),
            mode: Some(2),
            autoplay_transition_delay: Some(1.25),
            autoplay_build_delay: Some(0.75),
            idle_timer_active: Some(true),
            idle_timer_delay: Some(30.0),
            soundtrack: Some(reference(15)),
            automatically_plays_upon_open: Some(false),
            slide_list: Some(reference(16)),
        }
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "Each emitted varint byte intentionally contains only the low seven bits."
    )]
    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 0x80 {
            output.push((value as u8 & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn push_overlong_varint(value: u64, output: &mut Vec<u8>) {
        let start = output.len();
        push_varint(value, output);
        if let Some(last) = output.last_mut() {
            *last |= 0x80;
        }
        output.push(0);
        debug_assert!(output.len() >= start + 2);
    }

    fn opaque_overlong_varint_field(field_number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_overlong_varint(u64::from(field_number) << 3, &mut output);
        push_overlong_varint(value, &mut output);
        output
    }

    fn opaque_overlong_length_delimited_field(field_number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_overlong_varint((u64::from(field_number) << 3) | 2, &mut output);
        push_overlong_varint(payload.len() as u64, &mut output);
        output.extend_from_slice(payload);
        output
    }

    fn varint_field(field_number: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(u64::from(field_number) << 3, &mut output);
        push_varint(value, &mut output);
        output
    }

    fn fixed32_field(field_number: u32, value: f32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(field_number) << 3) | 5, &mut output);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn fixed64_field(field_number: u32, value: f64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint((u64::from(field_number) << 3) | 1, &mut output);
        output.extend_from_slice(&value.to_le_bytes());
        output
    }

    fn length_delimited_field(field_number: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::with_capacity(payload.len() + 8);
        push_varint((u64::from(field_number) << 3) | 2, &mut output);
        push_varint(payload.len() as u64, &mut output);
        output.extend_from_slice(payload);
        output
    }

    fn minimal_show(slides: &[u64]) -> Vec<u8> {
        let mut tree = Vec::new();
        for identifier in slides {
            tree.extend(length_delimited_field(
                SLIDE_TREE_SLIDES_FIELD,
                &varint_field(REFERENCE_IDENTIFIER_FIELD, *identifier),
            ));
        }
        minimal_show_with_tree(&tree)
    }

    fn minimal_show_with_tree(tree: &[u8]) -> Vec<u8> {
        let mut source = length_delimited_field(
            SHOW_THEME_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 1),
        );
        source.extend(length_delimited_field(SHOW_SLIDE_TREE_FIELD, tree));
        let mut size = fixed32_field(SIZE_WIDTH_FIELD, 1024.0);
        size.extend(fixed32_field(SIZE_HEIGHT_FIELD, 768.0));
        source.extend(length_delimited_field(SHOW_SIZE_FIELD, &size));
        source.extend(length_delimited_field(
            SHOW_STYLESHEET_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 2),
        ));
        source
    }

    fn assert_error<T>(result: Result<T, DecodeError>) -> DecodeError {
        match result {
            Err(error) => error,
            Ok(_) => panic!("malformed Keynote show unexpectedly decoded"),
        }
    }

    #[test]
    fn canonical_prost_show_matches_projection() -> Result<(), Box<dyn std::error::Error>> {
        let expected = show(&[30, 10, 20]);
        let source = expected.encode_to_vec();
        let native = kn::ShowArchive::decode(source.as_slice())?;
        let snapshot = decode_show(&source, options(&source, 3))?;
        let settings_only = decode_settings(&source, options(&source, 3))?;

        assert_eq!(
            snapshot.slide_node_identifiers(),
            native
                .slide_tree
                .slides
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>()
        );
        assert_eq!(
            snapshot.size().width().to_bits(),
            native.size.width.to_bits()
        );
        assert_eq!(
            snapshot.size().height().to_bits(),
            native.size.height.to_bits()
        );
        assert_eq!(settings_only.size(), snapshot.size());
        assert_eq!(settings_only.raw_settings(), snapshot.raw_settings());
        let settings = snapshot.raw_settings();
        assert_eq!(
            settings.slide_numbers_visible(),
            native.slide_numbers_visible
        );
        assert_eq!(settings.loop_presentation(), native.loop_presentation);
        assert_eq!(settings.mode_raw(), native.mode);
        assert_eq!(
            settings.autoplay_transition_delay().map(f64::to_bits),
            native.autoplay_transition_delay.map(f64::to_bits)
        );
        assert_eq!(
            settings.autoplay_build_delay().map(f64::to_bits),
            native.autoplay_build_delay.map(f64::to_bits)
        );
        assert_eq!(settings.idle_timer_active(), native.idle_timer_active);
        assert_eq!(
            settings.idle_timer_delay().map(f64::to_bits),
            native.idle_timer_delay.map(f64::to_bits)
        );
        assert_eq!(
            settings.automatically_plays_upon_open(),
            native.automatically_plays_upon_open
        );
        assert!(snapshot.has_deprecated_root_slide_node());
        assert!(snapshot.has_slide_list());
        Ok(())
    }

    #[test]
    fn native_show_oracle_preserves_exact_scalars_and_order()
    -> Result<(), Box<dyn std::error::Error>> {
        let snapshot = decode_show(&NATIVE_SHOW, options(&NATIVE_SHOW, 1))?;
        let native = kn::ShowArchive::decode(NATIVE_SHOW.as_slice())?;
        assert_eq!(snapshot.slide_node_identifiers(), [2_652_149]);
        assert_eq!(
            native.soundtrack.map(|reference| reference.identifier),
            Some(2_651_093)
        );
        assert_eq!(snapshot.size().width().to_bits(), 0x44f0_0000);
        assert_eq!(snapshot.size().height().to_bits(), 0x4487_0000);
        let settings = snapshot.raw_settings();
        assert_eq!(settings.loop_presentation(), Some(false));
        assert_eq!(settings.mode_raw(), Some(0));
        assert_eq!(
            settings.autoplay_transition_delay().map(f64::to_bits),
            Some(5.0f64.to_bits())
        );
        assert_eq!(
            settings.autoplay_build_delay().map(f64::to_bits),
            Some(2.0f64.to_bits())
        );
        assert_eq!(settings.idle_timer_active(), Some(false));
        assert_eq!(
            settings.idle_timer_delay().map(f64::to_bits),
            Some(900.0f64.to_bits())
        );
        assert_eq!(settings.automatically_plays_upon_open(), Some(false));
        assert!(!snapshot.has_deprecated_root_slide_node());
        assert!(!snapshot.has_slide_list());
        Ok(())
    }

    #[test]
    fn equal_and_zero_slide_identifiers_remain_schema_faithful()
    -> Result<(), Box<dyn std::error::Error>> {
        let source = minimal_show(&[0, 7, 7]);
        assert_eq!(
            decode_show(&source, options(&source, 3))?.slide_node_identifiers(),
            [0, 7, 7]
        );
        Ok(())
    }

    #[test]
    fn canonical_negative_int32_values_are_accepted() -> Result<(), Box<dyn std::error::Error>> {
        for (encoded, expected) in [(u64::MAX, -1), (MIN_SIGN_EXTENDED_INT32, i32::MIN)] {
            let mut source = minimal_show(&[]);
            source.extend(varint_field(SHOW_MODE_FIELD, encoded));
            assert_eq!(
                decode_show(&source, options(&source, 0))?
                    .raw_settings()
                    .mode_raw(),
                Some(expected)
            );
        }

        let mut slide_reference = varint_field(REFERENCE_IDENTIFIER_FIELD, 4);
        slide_reference.extend(varint_field(REFERENCE_DEPRECATED_TYPE_FIELD, u64::MAX));
        let tree = length_delimited_field(SLIDE_TREE_SLIDES_FIELD, &slide_reference);
        let source = minimal_show_with_tree(&tree);
        assert_eq!(
            decode_show(&source, options(&source, 1))?.slide_node_identifiers(),
            [4]
        );
        Ok(())
    }

    #[test]
    fn topology_flags_are_content_free_and_force_references() {
        let flat = minimal_show(&[4, 5]);
        let flat_snapshot = decode_show(&flat, options(&flat, 2)).unwrap_or_else(|error| {
            panic!("flat topology must decode: {error}");
        });
        assert!(!flat_snapshot.has_deprecated_root_slide_node());
        assert!(!flat_snapshot.has_slide_list());

        let invalid_root = length_delimited_field(SLIDE_TREE_ROOT_FIELD, &[]);
        let invalid_root_show = minimal_show_with_tree(&invalid_root);
        let root_error = assert_error(decode_show(
            &invalid_root_show,
            options(&invalid_root_show, 0),
        ));
        assert_eq!(
            root_error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );

        let mut invalid_list = minimal_show(&[]);
        invalid_list.extend(length_delimited_field(SHOW_SLIDE_LIST_FIELD, &[]));
        let list_error = assert_error(decode_show(&invalid_list, options(&invalid_list, 0)));
        assert_eq!(
            list_error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn slide_reference_limit_is_inclusive_and_typed() -> Result<(), Box<dyn std::error::Error>> {
        let source = minimal_show(&[1, 2, 3]);
        assert_eq!(
            decode_show(&source, options(&source, 3))?.slide_node_identifiers(),
            [1, 2, 3]
        );
        let error = assert_error(decode_show(&source, options(&source, 2)));
        assert_eq!(error.slide_reference_limit_values(), Some((3, 2)));
        Ok(())
    }

    #[test]
    fn every_known_singular_scope_rejects_duplicates() {
        for (field_number, field_name, payload) in [
            (
                SHOW_THEME_FIELD,
                "KN.ShowArchive.theme",
                varint_field(REFERENCE_IDENTIFIER_FIELD, 9),
            ),
            (
                SHOW_SLIDE_TREE_FIELD,
                "KN.ShowArchive.slideTree",
                Vec::new(),
            ),
            (SHOW_SIZE_FIELD, "KN.ShowArchive.size", {
                let mut value = fixed32_field(SIZE_WIDTH_FIELD, 1.0);
                value.extend(fixed32_field(SIZE_HEIGHT_FIELD, 1.0));
                value
            }),
            (
                SHOW_STYLESHEET_FIELD,
                "KN.ShowArchive.stylesheet",
                varint_field(REFERENCE_IDENTIFIER_FIELD, 9),
            ),
        ] {
            let mut duplicate = minimal_show(&[]);
            duplicate.extend(length_delimited_field(field_number, &payload));
            let error = assert_error(decode_show(&duplicate, options(&duplicate, 0)));
            assert_eq!(error.duplicate_singular_field(), Some(field_name));
        }

        let mut duplicate_setting = minimal_show(&[]);
        duplicate_setting.extend(varint_field(SHOW_LOOP_PRESENTATION_FIELD, 0));
        duplicate_setting.extend(varint_field(SHOW_LOOP_PRESENTATION_FIELD, 1));
        let setting_error = assert_error(decode_show(
            &duplicate_setting,
            options(&duplicate_setting, 0),
        ));
        assert_eq!(
            setting_error.duplicate_singular_field(),
            Some("KN.ShowArchive.loop_presentation")
        );

        let reference_payload = varint_field(REFERENCE_IDENTIFIER_FIELD, 9);
        for (field_name, encoded_field) in [
            (
                "KN.ShowArchive.uiState",
                length_delimited_field(SHOW_UI_STATE_FIELD, &reference_payload),
            ),
            (
                "KN.ShowArchive.slideNumbersVisible",
                varint_field(SHOW_SLIDE_NUMBERS_VISIBLE_FIELD, 1),
            ),
            (
                "KN.ShowArchive.recording",
                length_delimited_field(SHOW_RECORDING_FIELD, &reference_payload),
            ),
            ("KN.ShowArchive.mode", varint_field(SHOW_MODE_FIELD, 0)),
            (
                "KN.ShowArchive.autoplay_transition_delay",
                fixed64_field(SHOW_AUTOPLAY_TRANSITION_DELAY_FIELD, 5.0),
            ),
            (
                "KN.ShowArchive.autoplay_build_delay",
                fixed64_field(SHOW_AUTOPLAY_BUILD_DELAY_FIELD, 2.0),
            ),
            (
                "KN.ShowArchive.idle_timer_active",
                varint_field(SHOW_IDLE_TIMER_ACTIVE_FIELD, 0),
            ),
            (
                "KN.ShowArchive.idle_timer_delay",
                fixed64_field(SHOW_IDLE_TIMER_DELAY_FIELD, 900.0),
            ),
            (
                "KN.ShowArchive.soundtrack",
                length_delimited_field(SHOW_SOUNDTRACK_FIELD, &reference_payload),
            ),
            (
                "KN.ShowArchive.automatically_plays_upon_open",
                varint_field(SHOW_AUTOMATICALLY_PLAYS_UPON_OPEN_FIELD, 0),
            ),
            (
                "KN.ShowArchive.slideList",
                length_delimited_field(SHOW_SLIDE_LIST_FIELD, &reference_payload),
            ),
        ] {
            let mut duplicate = minimal_show(&[]);
            duplicate.extend_from_slice(&encoded_field);
            duplicate.extend_from_slice(&encoded_field);
            let error = assert_error(decode_show(&duplicate, options(&duplicate, 0)));
            assert_eq!(error.duplicate_singular_field(), Some(field_name));
        }

        let mut duplicate_identifier = varint_field(REFERENCE_IDENTIFIER_FIELD, 1);
        duplicate_identifier.extend(varint_field(REFERENCE_IDENTIFIER_FIELD, 2));
        let tree = length_delimited_field(SLIDE_TREE_SLIDES_FIELD, &duplicate_identifier);
        let duplicate_reference_show = minimal_show_with_tree(&tree);
        let reference_error = assert_error(decode_show(
            &duplicate_reference_show,
            options(&duplicate_reference_show, 1),
        ));
        assert_eq!(
            reference_error.duplicate_singular_field(),
            Some("TSP.Reference.identifier")
        );

        for (field_number, field_name, value) in [
            (
                REFERENCE_DEPRECATED_TYPE_FIELD,
                "TSP.Reference.deprecated_type",
                1,
            ),
            (
                REFERENCE_DEPRECATED_EXTERNAL_FIELD,
                "TSP.Reference.deprecated_is_external",
                0,
            ),
        ] {
            let mut duplicate_nested = varint_field(REFERENCE_IDENTIFIER_FIELD, 1);
            duplicate_nested.extend(varint_field(field_number, value));
            duplicate_nested.extend(varint_field(field_number, value));
            let nested_tree = length_delimited_field(SLIDE_TREE_SLIDES_FIELD, &duplicate_nested);
            let nested_show = minimal_show_with_tree(&nested_tree);
            let nested_error = assert_error(decode_show(&nested_show, options(&nested_show, 1)));
            assert_eq!(nested_error.duplicate_singular_field(), Some(field_name));
        }

        let mut duplicate_size = fixed32_field(SIZE_WIDTH_FIELD, 1.0);
        duplicate_size.extend(fixed32_field(SIZE_WIDTH_FIELD, 2.0));
        duplicate_size.extend(fixed32_field(SIZE_HEIGHT_FIELD, 3.0));
        let mut duplicate_size_show = length_delimited_field(
            SHOW_THEME_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 1),
        );
        duplicate_size_show.extend(length_delimited_field(SHOW_SLIDE_TREE_FIELD, &[]));
        duplicate_size_show.extend(length_delimited_field(SHOW_SIZE_FIELD, &duplicate_size));
        duplicate_size_show.extend(length_delimited_field(
            SHOW_STYLESHEET_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 2),
        ));
        let size_error = assert_error(decode_show(
            &duplicate_size_show,
            options(&duplicate_size_show, 0),
        ));
        assert_eq!(
            size_error.duplicate_singular_field(),
            Some("TSP.Size.width")
        );

        let root_reference = varint_field(REFERENCE_IDENTIFIER_FIELD, 3);
        let mut duplicate_root = length_delimited_field(SLIDE_TREE_ROOT_FIELD, &root_reference);
        duplicate_root.extend(length_delimited_field(
            SLIDE_TREE_ROOT_FIELD,
            &root_reference,
        ));
        let duplicate_root_show = minimal_show_with_tree(&duplicate_root);
        let root_error = assert_error(decode_show(
            &duplicate_root_show,
            options(&duplicate_root_show, 0),
        ));
        assert_eq!(
            root_error.duplicate_singular_field(),
            Some("KN.SlideTreeArchive.rootSlideNode")
        );
    }

    #[test]
    fn strict_preflight_rejects_noncanonical_known_wire() {
        let mut bad_bool = minimal_show(&[]);
        bad_bool.extend(varint_field(SHOW_LOOP_PRESENTATION_FIELD, 2));
        let bool_error = assert_error(decode_show(&bad_bool, options(&bad_bool, 0)));
        assert_eq!(
            bool_error.noncanonical_reason(),
            Some("bool scalar is not zero or one")
        );

        let mut bad_int32 = minimal_show(&[]);
        bad_int32.extend(varint_field(SHOW_MODE_FIELD, 0xffff_ffff));
        let int32_error = assert_error(decode_show(&bad_int32, options(&bad_int32, 0)));
        assert_eq!(
            int32_error.noncanonical_reason(),
            Some("int32 scalar is not a sign-extended 32-bit value")
        );

        let canonical = minimal_show(&[]);
        let mut overlong_key = Vec::with_capacity(canonical.len() + 1);
        overlong_key.extend([0x92, 0x00]);
        overlong_key.extend_from_slice(&canonical[1..]);
        let key_error = assert_error(decode_show(
            &overlong_key,
            DecodeOptions::new(overlong_key.len(), 0, 8),
        ));
        assert_eq!(key_error.noncanonical_reason(), Some("protobuf field key"));

        let mut overlong_length = Vec::with_capacity(canonical.len() + 1);
        overlong_length.push(canonical[0]);
        overlong_length.extend([canonical[1] | 0x80, 0x00]);
        overlong_length.extend_from_slice(&canonical[2..]);
        let length_error = assert_error(decode_show(
            &overlong_length,
            DecodeOptions::new(overlong_length.len(), 0, 8),
        ));
        assert_eq!(
            length_error.noncanonical_reason(),
            Some("length-delimited size")
        );

        let overlong_identifier = [0x08, 0x81, 0x00];
        let overlong_tree = length_delimited_field(SLIDE_TREE_SLIDES_FIELD, &overlong_identifier);
        let overlong_reference_show = minimal_show_with_tree(&overlong_tree);
        let value_error = assert_error(decode_show(
            &overlong_reference_show,
            options(&overlong_reference_show, 1),
        ));
        assert_eq!(
            value_error.noncanonical_reason(),
            Some("protobuf varint value")
        );
    }

    #[test]
    fn opaque_unknown_fields_keep_permissive_framing_in_every_known_scope()
    -> Result<(), Box<dyn std::error::Error>> {
        let unknown_varint = opaque_overlong_varint_field(100, 1);
        let unknown_bytes = opaque_overlong_length_delimited_field(101, b"opaque");

        let mut show_unknown = minimal_show(&[9]);
        show_unknown.extend_from_slice(&unknown_varint);
        show_unknown.extend_from_slice(&unknown_bytes);
        assert_eq!(
            decode_show(&show_unknown, options(&show_unknown, 1))?.slide_node_identifiers(),
            [9]
        );

        let mut reference = varint_field(REFERENCE_IDENTIFIER_FIELD, 7);
        reference.extend_from_slice(&unknown_varint);
        reference.extend_from_slice(&unknown_bytes);
        let mut tree = length_delimited_field(SLIDE_TREE_SLIDES_FIELD, &reference);
        tree.extend_from_slice(&unknown_varint);
        tree.extend_from_slice(&unknown_bytes);

        let mut size = fixed32_field(SIZE_WIDTH_FIELD, 1024.0);
        size.extend(fixed32_field(SIZE_HEIGHT_FIELD, 768.0));
        size.extend_from_slice(&unknown_varint);
        size.extend_from_slice(&unknown_bytes);
        let mut nested_unknown = length_delimited_field(
            SHOW_THEME_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 1),
        );
        nested_unknown.extend(length_delimited_field(SHOW_SLIDE_TREE_FIELD, &tree));
        nested_unknown.extend(length_delimited_field(SHOW_SIZE_FIELD, &size));
        nested_unknown.extend(length_delimited_field(
            SHOW_STYLESHEET_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 2),
        ));
        assert_eq!(
            decode_show(&nested_unknown, options(&nested_unknown, 1))?.slide_node_identifiers(),
            [7]
        );
        Ok(())
    }

    #[test]
    fn wide_slide_tree_streams_in_source_order_after_bounded_preflight()
    -> Result<(), Box<dyn std::error::Error>> {
        const SLIDES: usize = 4_096;

        let identifiers = (1..=SLIDES)
            .map(u64::try_from)
            .collect::<Result<Vec<_>, _>>()?;
        let source = minimal_show(&identifiers);
        let settings = decode_settings(&source, options(&source, SLIDES))?;
        assert_eq!(settings.size().width().to_bits(), 1_024.0_f32.to_bits());
        assert_eq!(settings.size().height().to_bits(), 768.0_f32.to_bits());
        let snapshot = decode_show(&source, options(&source, SLIDES))?;
        assert_eq!(snapshot.slide_node_identifiers(), identifiers);
        assert_eq!(
            assert_error(decode_settings(&source, options(&source, SLIDES - 1)))
                .slide_reference_limit_values(),
            Some((SLIDES, SLIDES - 1))
        );
        Ok(())
    }

    #[test]
    fn required_nested_fields_and_projected_wire_types_are_enforced() {
        let empty_error = assert_error(decode_show(&[], DecodeOptions::new(1, 0, 8)));
        assert_eq!(
            empty_error.missing_required_field(),
            Some("KN.ShowArchive.theme")
        );

        let mut size_missing_height = length_delimited_field(
            SHOW_THEME_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 1),
        );
        size_missing_height.extend(length_delimited_field(SHOW_SLIDE_TREE_FIELD, &[]));
        size_missing_height.extend(length_delimited_field(
            SHOW_SIZE_FIELD,
            &fixed32_field(SIZE_WIDTH_FIELD, 1.0),
        ));
        size_missing_height.extend(length_delimited_field(
            SHOW_STYLESHEET_FIELD,
            &varint_field(REFERENCE_IDENTIFIER_FIELD, 2),
        ));
        let size_error = assert_error(decode_show(
            &size_missing_height,
            options(&size_missing_height, 0),
        ));
        assert_eq!(size_error.missing_required_field(), Some("TSP.Size.height"));

        let mut scalar = minimal_show(&[]);
        scalar.extend(length_delimited_field(
            SHOW_SLIDE_NUMBERS_VISIBLE_FIELD,
            &[],
        ));
        assert!(decode_show(&scalar, options(&scalar, 0)).is_err());

        let malformed_tree = varint_field(SLIDE_TREE_SLIDES_FIELD, 1);
        let slide = minimal_show_with_tree(&malformed_tree);
        assert!(decode_show(&slide, options(&slide, 1)).is_err());

        let mut reference = minimal_show(&[]);
        reference.extend(length_delimited_field(
            SHOW_RECORDING_FIELD,
            &length_delimited_field(REFERENCE_IDENTIFIER_FIELD, &[]),
        ));
        assert!(decode_show(&reference, options(&reference, 0)).is_err());
    }

    #[test]
    fn finite_message_and_recursion_limits_are_enforced() {
        let source = minimal_show(&[1]);
        assert!(decode_show(&source, DecodeOptions::new(source.len() - 1, 1, 8)).is_err());
        assert!(decode_show(&source, DecodeOptions::new(source.len(), 1, 1)).is_err());
        assert!(decode_show(&source, DecodeOptions::new(usize::MAX, 1, 8)).is_err());
        assert!(decode_show(&source, DecodeOptions::new(source.len(), 1, 65)).is_err());
    }

    #[test]
    fn exact_full_show_resource_boundaries_are_typed_and_inclusive()
    -> Result<(), Box<dyn std::error::Error>> {
        const EXACT_FIELDS: usize = 12;
        const EXACT_WORK: usize = 94;

        let source = minimal_show(&[7]);
        assert_eq!(source.len(), 26);
        let exact = DecodeOptions::new(source.len(), 1, 3)
            .with_max_fields(EXACT_FIELDS)
            .with_max_work_bytes(EXACT_WORK);
        assert_eq!(decode_show(&source, exact)?.slide_node_identifiers(), [7]);

        let bytes = assert_error(decode_show(
            &source,
            DecodeOptions::new(source.len() - 1, 1, 3)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK),
        ));
        assert_eq!(
            bytes.wire_resource_limit(),
            Some(WireResourceLimit::Bytes {
                observed: source.len(),
                maximum: source.len() - 1,
            })
        );

        let fields = assert_error(decode_show(
            &source,
            DecodeOptions::new(source.len(), 1, 3)
                .with_max_fields(EXACT_FIELDS - 1)
                .with_max_work_bytes(EXACT_WORK),
        ));
        assert_eq!(
            fields.field_limit_values(),
            Some((EXACT_FIELDS, EXACT_FIELDS - 1))
        );

        let work = assert_error(decode_show(
            &source,
            DecodeOptions::new(source.len(), 1, 3)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK - 1),
        ));
        assert_eq!(work.work_limit_values(), Some((EXACT_WORK, EXACT_WORK - 1)));

        let nesting = assert_error(decode_show(
            &source,
            DecodeOptions::new(source.len(), 1, 2)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK),
        ));
        assert_eq!(
            nesting.wire_resource_limit(),
            Some(WireResourceLimit::Nesting {
                observed: 3,
                maximum: 2,
            })
        );
        Ok(())
    }

    #[test]
    fn settings_only_has_an_exact_smaller_aggregate_budget()
    -> Result<(), Box<dyn std::error::Error>> {
        const EXACT_FIELDS: usize = 10;
        const EXACT_WORK: usize = 86;

        let source = minimal_show(&[7]);
        let exact = DecodeOptions::new(source.len(), 1, 3)
            .with_max_fields(EXACT_FIELDS)
            .with_max_work_bytes(EXACT_WORK);
        assert_eq!(
            decode_settings(&source, exact)?.size().width().to_bits(),
            1024.0f32.to_bits()
        );

        let fields = assert_error(decode_settings(
            &source,
            DecodeOptions::new(source.len(), 1, 3)
                .with_max_fields(EXACT_FIELDS - 1)
                .with_max_work_bytes(EXACT_WORK),
        ));
        assert_eq!(
            fields.field_limit_values(),
            Some((EXACT_FIELDS, EXACT_FIELDS - 1))
        );

        let work = assert_error(decode_settings(
            &source,
            DecodeOptions::new(source.len(), 1, 3)
                .with_max_fields(EXACT_FIELDS)
                .with_max_work_bytes(EXACT_WORK - 1),
        ));
        assert_eq!(work.work_limit_values(), Some((EXACT_WORK, EXACT_WORK - 1)));
        Ok(())
    }

    #[test]
    fn invalid_resource_profiles_and_structural_wire_adversaries_are_typed() {
        let source = minimal_show(&[]);
        for (configured, expected) in [(0, 0), (65, 65)] {
            let error = assert_error(decode_settings(
                &source,
                DecodeOptions::new(source.len(), 0, configured),
            ));
            assert_eq!(
                error.wire_resource_limit(),
                Some(WireResourceLimit::Nesting {
                    observed: expected,
                    maximum: 64,
                })
            );
        }

        let malformed: [&[u8]; 8] = [
            &[0x80],
            &[0x00],
            &[0x0f],
            &[0x1a, 0x02, 0x12],
            &[0x0b],
            &[0x0c],
            &[0x12, 0x01, 0x80],
            &[
                0x08, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x02,
            ],
        ];
        for bytes in malformed {
            assert!(decode_settings(bytes, DecodeOptions::new(bytes.len().max(1), 0, 8),).is_err());
        }
    }

    #[test]
    fn large_opaque_unknown_field_is_not_retained_or_projected()
    -> Result<(), Box<dyn std::error::Error>> {
        let mut source = minimal_show(&[9]);
        source.extend(length_delimited_field(100, &vec![0xa5; 256 * 1024]));
        let snapshot = decode_show(&source, options(&source, 1))?;
        assert_eq!(snapshot.slide_node_identifiers(), [9]);
        Ok(())
    }
}
