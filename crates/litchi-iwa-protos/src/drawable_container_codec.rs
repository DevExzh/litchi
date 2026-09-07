//! Strict lazy projection for native TSD container and group edges.
//!
//! Object-index construction needs only the parent and child references from
//! `TSD.ContainerArchive` and `TSD.GroupArchive`.  This module keeps the
//! caller-owned archive bytes authoritative: it performs a complete bounded
//! wire scan before asking Buffa for borrowed lazy views of the selected
//! fields, and publishes only non-zero object identifiers.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The strict scanner intentionally precedes the Buffa parity pass."
)]

use std::{fmt, num::NonZeroU64};

use buffa::DecodeOptions as BuffaDecodeOptions;

use crate::buffa_drawable_container_generated::LitchiIwaProjection as projection;

const CONTAINER_PARENT_FIELD: u32 = 2;
const CONTAINER_CHILDREN_FIELD: u32 = 3;
const GROUP_SUPER_FIELD: u32 = 1;
const GROUP_CHILDREN_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const REFERENCE_DEPRECATED_TYPE_FIELD: u32 = 2;
const REFERENCE_DEPRECATED_EXTERNAL_FIELD: u32 = 3;
const MAX_RECURSION_LIMIT: u32 = 64;
const MIN_SIGN_EXTENDED_INT32: u64 = 0xffff_ffff_8000_0000;

/// Finite resource policy for one native container or group payload.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct DecodeOptions {
    max_message_bytes: usize,
    max_fields: usize,
    max_work_bytes: usize,
    recursion_limit: u32,
    recursion_maximum: u32,
    max_references: usize,
}

impl DecodeOptions {
    /// Build an explicit finite bytes/fields/work/nesting policy.
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
            recursion_maximum: recursion_limit,
            max_references: max_fields,
        }
    }

    /// Derive a conservative finite policy from one caller-owned payload.
    #[must_use]
    pub fn for_source(source: &[u8]) -> Self {
        let bytes = source.len().max(1);
        Self {
            max_message_bytes: bytes,
            max_fields: bytes.saturating_mul(8).max(1),
            max_work_bytes: bytes.saturating_mul(16).max(1),
            recursion_limit: 8,
            recursion_maximum: 8,
            max_references: bytes,
        }
    }

    /// Replace the strict field-visit ceiling.
    #[must_use]
    pub const fn with_max_fields(mut self, maximum: usize) -> Self {
        self.max_fields = maximum;
        self
    }

    /// Replace the strict-plus-Buffa work ceiling.
    #[must_use]
    pub const fn with_max_work_bytes(mut self, maximum: usize) -> Self {
        self.max_work_bytes = maximum;
        self
    }

    /// Replace the protobuf nesting ceiling.
    #[must_use]
    pub const fn with_recursion_limit(mut self, maximum: u32) -> Self {
        self.recursion_limit = maximum;
        self.recursion_maximum = maximum;
        self
    }

    /// Replace the number of published parent/child references permitted.
    #[must_use]
    pub const fn with_max_references(mut self, maximum: usize) -> Self {
        self.max_references = maximum;
        self
    }

    fn buffa(self) -> BuffaDecodeOptions {
        BuffaDecodeOptions::new()
            .with_max_message_size(self.max_message_bytes)
            .with_unknown_field_limit(0)
            .with_element_memory_limit(0)
            .with_recursion_limit(self.recursion_limit)
    }

    fn descend(self) -> Result<Self, DecodeError> {
        let recursion_limit = self.recursion_limit.checked_sub(1).ok_or_else(|| {
            DecodeError::recursion_limit(
                self.recursion_maximum.saturating_add(1),
                self.recursion_maximum,
            )
        })?;
        Ok(Self {
            recursion_limit,
            ..self
        })
    }
}

/// Validated object-index edges selected from one native container or group.
///
/// The parent, when present, is yielded first, followed by the archive's
/// children in wire order.  IDs are compact `NonZeroU64` values, so callers
/// never need to handle malformed zero references or raw protobuf objects.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ReferenceSnapshot {
    parent: Option<NonZeroU64>,
    children: Vec<NonZeroU64>,
}

impl ReferenceSnapshot {
    /// Return the optional parent edge.
    #[must_use]
    pub const fn parent(&self) -> Option<NonZeroU64> {
        self.parent
    }

    /// Return child edges in their encoded order.
    #[must_use]
    pub fn children(&self) -> impl ExactSizeIterator<Item = NonZeroU64> + '_ {
        self.children.iter().copied()
    }

    /// Return parent first, followed by all children in encoded order.
    pub fn references(&self) -> impl Iterator<Item = NonZeroU64> + '_ {
        self.parent.into_iter().chain(self.children.iter().copied())
    }

    /// Split the compact snapshot into its optional parent and child edges.
    #[must_use]
    pub fn into_parts(self) -> (Option<NonZeroU64>, Vec<NonZeroU64>) {
        (self.parent, self.children)
    }
}

/// Resource axis rejected by strict container/group ingress.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum DecodeLimit {
    /// Input bytes exceeded the configured ceiling.
    Bytes { observed: usize, maximum: usize },
    /// Encoded fields exceeded the configured ceiling.
    Fields { observed: usize, maximum: usize },
    /// Strict and lazy-view work exceeded the configured ceiling.
    Work { observed: usize, maximum: usize },
    /// Protobuf nesting exceeded the configured ceiling.
    Nesting { observed: u32, maximum: u32 },
    /// Published parent/child references exceeded the configured ceiling.
    References { observed: usize, maximum: usize },
}

/// Failure from strict container/group decoding or its Buffa parity check.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DecodeError {
    kind: DecodeErrorKind,
}

#[derive(Debug, Clone, PartialEq, Eq)]
enum DecodeErrorKind {
    Wire(buffa::DecodeError),
    Limit(DecodeLimit),
    MissingRequired(&'static str),
    DuplicateSingular(&'static str),
    NonCanonical(&'static str),
    Allocation(&'static str),
    Projection,
}

impl DecodeError {
    const fn limit(limit: DecodeLimit) -> Self {
        Self {
            kind: DecodeErrorKind::Limit(limit),
        }
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

    const fn allocation(resource: &'static str) -> Self {
        Self {
            kind: DecodeErrorKind::Allocation(resource),
        }
    }

    const fn projection() -> Self {
        Self {
            kind: DecodeErrorKind::Projection,
        }
    }

    const fn recursion_limit(observed: u32, maximum: u32) -> Self {
        Self::limit(DecodeLimit::Nesting { observed, maximum })
    }

    /// Return the rejected resource, if this is a finite-limit failure.
    #[must_use]
    pub const fn resource_limit(&self) -> Option<DecodeLimit> {
        match self.kind {
            DecodeErrorKind::Limit(limit) => Some(limit),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Allocation(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the missing required field, if applicable.
    #[must_use]
    pub const fn missing_required_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::MissingRequired(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Allocation(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the duplicated singular field, if applicable.
    #[must_use]
    pub const fn duplicate_singular_field(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::DuplicateSingular(field) => Some(field),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::NonCanonical(_)
            | DecodeErrorKind::Allocation(_)
            | DecodeErrorKind::Projection => None,
        }
    }

    /// Return the stable non-canonical wire reason, if applicable.
    #[must_use]
    pub const fn noncanonical_reason(&self) -> Option<&'static str> {
        match self.kind {
            DecodeErrorKind::NonCanonical(reason) => Some(reason),
            DecodeErrorKind::Wire(_)
            | DecodeErrorKind::Limit(_)
            | DecodeErrorKind::MissingRequired(_)
            | DecodeErrorKind::DuplicateSingular(_)
            | DecodeErrorKind::Allocation(_)
            | DecodeErrorKind::Projection => None,
        }
    }
}

impl fmt::Display for DecodeError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match &self.kind {
            DecodeErrorKind::Wire(error) => error.fmt(formatter),
            DecodeErrorKind::Limit(DecodeLimit::Bytes { observed, maximum }) => write!(
                formatter,
                "TSD drawable-container projection byte limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Fields { observed, maximum }) => write!(
                formatter,
                "TSD drawable-container projection field limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Work { observed, maximum }) => write!(
                formatter,
                "TSD drawable-container projection work limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::Nesting { observed, maximum }) => write!(
                formatter,
                "TSD drawable-container projection nesting limit exceeded: {observed} > {maximum}"
            ),
            DecodeErrorKind::Limit(DecodeLimit::References { observed, maximum }) => write!(
                formatter,
                "TSD drawable-container projection reference limit exceeded: {observed} > {maximum}"
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
            DecodeErrorKind::Allocation(resource) => {
                write!(formatter, "allocation failed while reserving {resource}")
            },
            DecodeErrorKind::Projection => formatter.write_str(
                "TSD drawable-container strict preflight disagrees with the Buffa projection",
            ),
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

/// Decode parent/child edges from a native `TSD.ContainerArchive`.
pub fn decode_container_references(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ReferenceSnapshot, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_container(source, options, &mut budget)?;
    let view: projection::ContainerArchiveLazyView<'_> =
        options.buffa().decode_lazy_view(source)?;
    force_container(&view, &strict)?;
    strict.into_snapshot()
}

/// Decode parent/child edges from a native `TSD.GroupArchive`.
pub fn decode_group_references(
    source: &[u8],
    options: DecodeOptions,
) -> Result<ReferenceSnapshot, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(options);
    let strict = preflight_group(source, options, &mut budget)?;
    let view: projection::GroupArchiveLazyView<'_> = options.buffa().decode_lazy_view(source)?;
    force_group(&view, &strict)?;
    strict.into_snapshot()
}

fn validate_input(source: &[u8], options: DecodeOptions) -> Result<(), DecodeError> {
    let hard_bytes = usize::try_from(buffa::MAX_MESSAGE_BYTES)
        .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
    if options.max_message_bytes > hard_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: options.max_message_bytes,
            maximum: hard_bytes,
        }));
    }
    if source.len() > options.max_message_bytes {
        return Err(DecodeError::limit(DecodeLimit::Bytes {
            observed: source.len(),
            maximum: options.max_message_bytes,
        }));
    }
    if options.recursion_limit == 0 {
        return Err(DecodeError::recursion_limit(1, 0));
    }
    if options.recursion_limit > MAX_RECURSION_LIMIT {
        return Err(DecodeError::recursion_limit(
            options.recursion_limit,
            MAX_RECURSION_LIMIT,
        ));
    }
    Ok(())
}

#[derive(Debug)]
struct Budget {
    fields: usize,
    work_bytes: usize,
    references: usize,
    max_fields: usize,
    max_work_bytes: usize,
    max_references: usize,
}

impl Budget {
    const fn new(options: DecodeOptions) -> Self {
        Self {
            fields: 0,
            work_bytes: 0,
            references: 0,
            max_fields: options.max_fields,
            max_work_bytes: options.max_work_bytes,
            max_references: options.max_references,
        }
    }

    fn charge_field(&mut self) -> Result<(), DecodeError> {
        let observed = self.fields.saturating_add(1);
        if observed > self.max_fields {
            return Err(DecodeError::limit(DecodeLimit::Fields {
                observed,
                maximum: self.max_fields,
            }));
        }
        self.fields = observed;
        Ok(())
    }

    fn charge_work(&mut self, bytes: usize) -> Result<(), DecodeError> {
        let observed = self.work_bytes.saturating_add(bytes.saturating_mul(2));
        if observed > self.max_work_bytes {
            return Err(DecodeError::limit(DecodeLimit::Work {
                observed,
                maximum: self.max_work_bytes,
            }));
        }
        self.work_bytes = observed;
        Ok(())
    }

    fn charge_reference(&mut self) -> Result<(), DecodeError> {
        let observed = self.references.saturating_add(1);
        if observed > self.max_references {
            return Err(DecodeError::limit(DecodeLimit::References {
                observed,
                maximum: self.max_references,
            }));
        }
        self.references = observed;
        Ok(())
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct RawReference {
    identifier: u64,
    deprecated_type: Option<i32>,
    deprecated_is_external: Option<bool>,
}

#[derive(Debug, PartialEq, Eq)]
struct StrictSnapshot {
    parent: Option<RawReference>,
    children: Vec<RawReference>,
}

impl StrictSnapshot {
    fn into_snapshot(self) -> Result<ReferenceSnapshot, DecodeError> {
        let mut children = Vec::new();
        children
            .try_reserve(self.children.len())
            .map_err(|_allocation| DecodeError::allocation("container/group children"))?;
        for reference in self.children {
            if let Some(identifier) = NonZeroU64::new(reference.identifier) {
                children.push(identifier);
            }
        }
        Ok(ReferenceSnapshot {
            parent: self
                .parent
                .and_then(|reference| NonZeroU64::new(reference.identifier)),
            children,
        })
    }
}

fn preflight_container(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictSnapshot, DecodeError> {
    budget.charge_work(source.len())?;
    let mut parent = None;
    let mut children = Vec::new();
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
        match field.number {
            CONTAINER_PARENT_FIELD => {
                if parent.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSD.ContainerArchive.parent",
                    ));
                }
                budget.charge_reference()?;
                let payload = field.length_delimited()?;
                parent = Some(preflight_reference(payload, options.descend()?, budget)?);
            },
            CONTAINER_CHILDREN_FIELD => {
                budget.charge_reference()?;
                let payload = field.length_delimited()?;
                children
                    .try_reserve(1)
                    .map_err(|_allocation| DecodeError::allocation("container children"))?;
                children.push(preflight_reference(payload, options.descend()?, budget)?);
            },
            _ => {},
        }
    }
    Ok(StrictSnapshot { parent, children })
}

fn preflight_group(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<StrictSnapshot, DecodeError> {
    budget.charge_work(source.len())?;
    let mut super_parent = None;
    let mut super_seen = false;
    let mut children = Vec::new();
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
        match field.number {
            GROUP_SUPER_FIELD => {
                if super_seen {
                    return Err(DecodeError::duplicate_singular("TSD.GroupArchive.super"));
                }
                super_seen = true;
                let payload = field.length_delimited()?;
                super_parent = preflight_drawable(payload, options.descend()?, budget)?;
            },
            GROUP_CHILDREN_FIELD => {
                budget.charge_reference()?;
                let payload = field.length_delimited()?;
                children
                    .try_reserve(1)
                    .map_err(|_allocation| DecodeError::allocation("group children"))?;
                children.push(preflight_reference(payload, options.descend()?, budget)?);
            },
            _ => {},
        }
    }
    if !super_seen {
        return Err(DecodeError::missing_required("TSD.GroupArchive.super"));
    }
    Ok(StrictSnapshot {
        parent: super_parent,
        children,
    })
}

fn preflight_drawable(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<Option<RawReference>, DecodeError> {
    budget.charge_work(source.len())?;
    let mut parent = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
        if field.number != DRAWABLE_PARENT_FIELD {
            continue;
        }
        if parent.is_some() {
            return Err(DecodeError::duplicate_singular(
                "TSD.DrawableArchive.parent",
            ));
        }
        budget.charge_reference()?;
        let payload = field.length_delimited()?;
        parent = Some(preflight_reference(payload, options.descend()?, budget)?);
    }
    Ok(parent)
}

fn preflight_reference(
    source: &[u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<RawReference, DecodeError> {
    budget.charge_work(source.len())?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut deprecated_is_external = None;
    let mut remaining = source;
    while let Some(field) = next_strict_field(
        &mut remaining,
        options.recursion_limit,
        options.recursion_maximum,
        budget,
    )? {
        match field.number {
            REFERENCE_IDENTIFIER_FIELD => {
                if identifier.is_some() {
                    return Err(DecodeError::duplicate_singular("TSP.Reference.identifier"));
                }
                let value = field.varint()?;
                identifier = Some(value);
            },
            REFERENCE_DEPRECATED_TYPE_FIELD => {
                if deprecated_type.is_some() {
                    return Err(DecodeError::duplicate_singular(
                        "TSP.Reference.deprecated_type",
                    ));
                }
                deprecated_type = Some(require_canonical_int32(field.varint()?)?);
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
    Ok(RawReference {
        identifier: identifier
            .ok_or_else(|| DecodeError::missing_required("TSP.Reference.identifier"))?,
        deprecated_type,
        deprecated_is_external,
    })
}

fn force_container(
    view: &projection::ContainerArchiveLazyView<'_>,
    strict: &StrictSnapshot,
) -> Result<(), DecodeError> {
    let parent = view
        .parent
        .get()?
        .map(|reference| force_reference(&reference))
        .transpose()?;
    if parent != strict.parent {
        return Err(DecodeError::projection());
    }
    if view.children.len() != strict.children.len() {
        return Err(DecodeError::projection());
    }
    for (reference, expected) in view.children.iter().zip(strict.children.iter()) {
        let reference = reference?;
        if force_reference(&reference)? != *expected {
            return Err(DecodeError::projection());
        }
    }
    Ok(())
}

fn force_group(
    view: &projection::GroupArchiveLazyView<'_>,
    strict: &StrictSnapshot,
) -> Result<(), DecodeError> {
    let super_view = view
        .super_
        .get()?
        .ok_or_else(|| DecodeError::missing_required("TSD.GroupArchive.super"))?;
    let parent = super_view
        .parent
        .get()?
        .map(|reference| force_reference(&reference))
        .transpose()?;
    if parent != strict.parent {
        return Err(DecodeError::projection());
    }
    if view.children.len() != strict.children.len() {
        return Err(DecodeError::projection());
    }
    for (reference, expected) in view.children.iter().zip(strict.children.iter()) {
        let reference = reference?;
        if force_reference(&reference)? != *expected {
            return Err(DecodeError::projection());
        }
    }
    Ok(())
}

fn force_reference(view: &projection::ReferenceLazyView<'_>) -> Result<RawReference, DecodeError> {
    if !view.has_identifier() {
        return Err(DecodeError::missing_required("TSP.Reference.identifier"));
    }
    Ok(RawReference {
        identifier: view.identifier,
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

fn require_canonical_int32(value: u64) -> Result<i32, DecodeError> {
    if value <= u64::from(u32::MAX / 2) {
        return i32::try_from(value).map_err(|_conversion| DecodeError::projection());
    }
    if value >= MIN_SIGN_EXTENDED_INT32 {
        let truncated = u32::try_from(value & u64::from(u32::MAX))
            .map_err(|_conversion| DecodeError::projection())?;
        return Ok(i32::from_ne_bytes(truncated.to_ne_bytes()));
    }
    Err(DecodeError::noncanonical(
        "int32 scalar is not sign-extended",
    ))
}

#[derive(Clone, Copy, Debug)]
enum StrictValue<'source> {
    Varint(u64),
    Fixed64,
    LengthDelimited(&'source [u8]),
    Group,
    Fixed32,
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

    fn varint(self) -> Result<u64, DecodeError> {
        self.require_canonical_key()?;
        if self.wire_type != buffa::encoding::WireType::Varint {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: buffa::encoding::WireType::Varint as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("protobuf varint value"));
        }
        match self.value {
            StrictValue::Varint(value) => Ok(value),
            StrictValue::Fixed64
            | StrictValue::LengthDelimited(_)
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
        }
    }

    fn length_delimited(self) -> Result<&'source [u8], DecodeError> {
        self.require_canonical_key()?;
        if self.wire_type != buffa::encoding::WireType::LengthDelimited {
            return Err(buffa::DecodeError::WireTypeMismatch {
                field_number: self.number,
                expected: buffa::encoding::WireType::LengthDelimited as u8,
                actual: self.wire_type as u8,
            }
            .into());
        }
        if !self.canonical_value {
            return Err(DecodeError::noncanonical("length-delimited size"));
        }
        match self.value {
            StrictValue::LengthDelimited(value) => Ok(value),
            StrictValue::Varint(_)
            | StrictValue::Fixed64
            | StrictValue::Group
            | StrictValue::Fixed32 => Err(DecodeError::projection()),
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
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<Option<StrictField<'source>>, DecodeError> {
    match parse_strict_field(source, recursion_limit, recursion_maximum, budget)? {
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
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<Option<ParseItem<'source>>, DecodeError> {
    if source.is_empty() {
        return Ok(None);
    }
    let (encoded_tag, canonical_key) = take_varint(source)?;
    budget.charge_field()?;
    let raw_tag =
        u32::try_from(encoded_tag).map_err(|_conversion| buffa::DecodeError::InvalidFieldNumber)?;
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
            take_exact(source, 8)?;
            (StrictValue::Fixed64, true)
        },
        buffa::encoding::WireType::LengthDelimited => {
            let (encoded_length, canonical) = take_varint(source)?;
            let length = usize::try_from(encoded_length)
                .map_err(|_conversion| buffa::DecodeError::MessageTooLarge)?;
            let payload = take_exact(source, length)?;
            (StrictValue::LengthDelimited(payload), canonical)
        },
        buffa::encoding::WireType::StartGroup => {
            let child_limit = recursion_limit.checked_sub(1).ok_or_else(|| {
                DecodeError::recursion_limit(recursion_maximum.saturating_add(1), recursion_maximum)
            })?;
            skip_strict_group(source, field_number, child_limit, recursion_maximum, budget)?;
            (StrictValue::Group, true)
        },
        buffa::encoding::WireType::EndGroup => return Ok(Some(ParseItem::EndGroup(field_number))),
        buffa::encoding::WireType::Fixed32 => {
            take_exact(source, 4)?;
            (StrictValue::Fixed32, true)
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
    recursion_maximum: u32,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    loop {
        match parse_strict_field(source, recursion_limit, recursion_maximum, budget)? {
            Some(ParseItem::Field(_)) => {},
            Some(ParseItem::EndGroup(number)) if number == expected_field_number => return Ok(()),
            Some(ParseItem::EndGroup(number)) => {
                return Err(buffa::DecodeError::InvalidEndGroup(number).into());
            },
            None => return Err(buffa::DecodeError::UnexpectedEof.into()),
        }
    }
}

fn take_exact<'source>(
    source: &mut &'source [u8],
    length: usize,
) -> Result<&'source [u8], DecodeError> {
    if source.len() < length {
        return Err(buffa::DecodeError::UnexpectedEof.into());
    }
    let (value, rest) = source.split_at(length);
    *source = rest;
    Ok(value)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{tsd, tsp};
    use prost::Message as _;

    fn options(source: &[u8]) -> DecodeOptions {
        DecodeOptions::for_source(source)
            .with_max_fields(usize::MAX)
            .with_max_work_bytes(usize::MAX)
            .with_max_references(usize::MAX)
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(-7),
            deprecated_is_external: Some(false),
        }
    }

    fn varint_field(field: u32, value: u64) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, u64::from(field) << 3);
        push_varint(&mut output, value);
        output
    }

    fn length_field(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 2);
        push_varint(
            &mut output,
            u64::try_from(payload.len()).expect("fixture length fits u64"),
        );
        output.extend_from_slice(payload);
        output
    }

    fn start_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 3);
        output
    }

    fn end_group(field: u32) -> Vec<u8> {
        let mut output = Vec::new();
        push_varint(&mut output, (u64::from(field) << 3) | 4);
        output
    }

    fn push_varint(output: &mut Vec<u8>, mut value: u64) {
        loop {
            let mut byte = u8::try_from(value & 0x7f).expect("varint chunk fits u8");
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            output.push(byte);
            if value == 0 {
                return;
            }
        }
    }

    #[test]
    fn container_snapshot_matches_native_edges_and_skips_null_references() {
        let native = tsd::ContainerArchive {
            parent: Some(reference(7)),
            children: vec![reference(0), reference(21), reference(21), reference(22)],
            ..Default::default()
        };
        let source = native.encode_to_vec();
        let snapshot = decode_container_references(&source, options(&source)).expect("snapshot");

        assert_eq!(snapshot.parent().map(NonZeroU64::get), Some(7));
        assert_eq!(
            snapshot.children().map(NonZeroU64::get).collect::<Vec<_>>(),
            [21, 21, 22]
        );
        assert_eq!(
            snapshot
                .references()
                .map(NonZeroU64::get)
                .collect::<Vec<_>>(),
            [7, 21, 21, 22]
        );
        let decoded = tsd::ContainerArchive::decode(source.as_slice()).expect("native decode");
        assert_eq!(decoded.parent, native.parent);
        assert_eq!(decoded.children, native.children);
    }

    #[test]
    fn group_snapshot_matches_required_super_with_optional_parent() {
        let native = tsd::GroupArchive {
            super_: tsd::DrawableArchive::default(),
            children: vec![reference(0), reference(31)],
            ..Default::default()
        };
        let source = native.encode_to_vec();
        let snapshot = decode_group_references(&source, options(&source)).expect("snapshot");

        assert_eq!(snapshot.parent(), None);
        assert_eq!(
            snapshot
                .references()
                .map(NonZeroU64::get)
                .collect::<Vec<_>>(),
            [31]
        );
    }

    #[test]
    fn group_requires_super_but_not_super_parent() {
        let error = decode_group_references(&[], DecodeOptions::for_source(&[]))
            .expect_err("required group super");
        assert_eq!(
            error.missing_required_field(),
            Some("TSD.GroupArchive.super")
        );
    }

    #[test]
    fn malformed_late_child_is_rejected_before_snapshot_publication() {
        let valid = tsd::ContainerArchive {
            parent: Some(reference(7)),
            children: vec![reference(21)],
            ..Default::default()
        };
        let mut source = valid.encode_to_vec();
        source.extend(length_field(CONTAINER_CHILDREN_FIELD, &varint_field(2, 1)));

        let error = decode_container_references(&source, options(&source))
            .expect_err("missing identifier in late child");
        assert_eq!(
            error.missing_required_field(),
            Some("TSP.Reference.identifier")
        );
    }

    #[test]
    fn duplicate_parent_and_wrong_wire_are_rejected() {
        let native = tsd::ContainerArchive {
            parent: Some(reference(7)),
            ..Default::default()
        };
        let mut duplicate = native.encode_to_vec();
        duplicate.extend(length_field(
            CONTAINER_PARENT_FIELD,
            &reference(8).encode_to_vec(),
        ));
        let duplicate_error = decode_container_references(&duplicate, options(&duplicate))
            .expect_err("duplicate parent");
        assert_eq!(
            duplicate_error.duplicate_singular_field(),
            Some("TSD.ContainerArchive.parent")
        );

        let wrong_wire = varint_field(CONTAINER_PARENT_FIELD, 7);
        let wrong_wire_error = decode_container_references(&wrong_wire, options(&wrong_wire))
            .expect_err("wrong parent wire type");
        assert!(matches!(
            wrong_wire_error.kind,
            DecodeErrorKind::Wire(buffa::DecodeError::WireTypeMismatch { .. })
        ));
    }

    #[test]
    fn reference_varint_overflow_and_nonminimal_values_are_rejected() {
        let mut overflow_reference = vec![0x08];
        overflow_reference.extend([0x80; 9]);
        overflow_reference.push(0x02);
        let overflow_source = length_field(CONTAINER_CHILDREN_FIELD, &overflow_reference);
        let overflow_error =
            decode_container_references(&overflow_source, options(&overflow_source))
                .expect_err("overflowing identifier");
        assert!(matches!(
            overflow_error.kind,
            DecodeErrorKind::Wire(buffa::DecodeError::VarintTooLong)
        ));

        let nonminimal_reference = [0x08, 0x81, 0x00];
        let nonminimal_source = length_field(CONTAINER_CHILDREN_FIELD, &nonminimal_reference);
        let nonminimal_error =
            decode_container_references(&nonminimal_source, options(&nonminimal_source))
                .expect_err("non-minimal identifier");
        assert_eq!(
            nonminimal_error.noncanonical_reason(),
            Some("protobuf varint value")
        );
    }

    #[test]
    fn balanced_unknown_groups_are_scanned_and_recursive_groups_are_bounded() {
        let native = tsd::ContainerArchive {
            parent: Some(reference(7)),
            ..Default::default()
        };
        let mut valid_source = native.encode_to_vec();
        valid_source.extend(start_group(99));
        valid_source.extend(varint_field(100, 7));
        valid_source.extend(end_group(99));
        let snapshot = decode_container_references(&valid_source, options(&valid_source))
            .expect("balanced unknown group");
        assert_eq!(
            snapshot
                .references()
                .map(NonZeroU64::get)
                .collect::<Vec<_>>(),
            [7]
        );

        let recursive_source = [
            start_group(99),
            start_group(100),
            end_group(100),
            end_group(99),
        ]
        .concat();
        let error = decode_container_references(
            &recursive_source,
            options(&recursive_source).with_recursion_limit(1),
        )
        .expect_err("recursive unknown group");
        assert_eq!(
            error.resource_limit(),
            Some(DecodeLimit::Nesting {
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn finite_reference_field_and_work_limits_are_enforced() {
        let native = tsd::ContainerArchive {
            parent: Some(reference(7)),
            children: vec![reference(21)],
            ..Default::default()
        };
        let source = native.encode_to_vec();
        let reference_error =
            decode_container_references(&source, options(&source).with_max_references(1))
                .expect_err("reference cap");
        assert!(matches!(
            reference_error.resource_limit(),
            Some(DecodeLimit::References { .. })
        ));

        let field_error = decode_container_references(&source, options(&source).with_max_fields(1))
            .expect_err("field cap");
        assert!(matches!(
            field_error.resource_limit(),
            Some(DecodeLimit::Fields { .. })
        ));

        let work_error = decode_container_references(
            &source,
            options(&source).with_max_work_bytes(source.len()),
        )
        .expect_err("work cap");
        assert!(matches!(
            work_error.resource_limit(),
            Some(DecodeLimit::Work { .. })
        ));
    }
}
