//! The shape-style variation graph engine (record 0375a): the copy-on-write
//! `TSWP.ShapeStyleArchive` variations through which iWork gives one shape
//! or text box its own fill, stroke, shadow, effects, text layout or text
//! columns. It is the `litchi-iwa` host's `shapes::line_end::registry`
//! lifecycle (`direct_shape_style_overrides`, `shape_style_is_exclusive`,
//! the variation object, insert/replace/collapse and the stylesheet
//! registry), ported operation for operation.
//!
//! Two layers:
//!
//! - **Plan** (payload-only, for a leaf's own transaction machinery): the
//!   leaf resolves style payloads through [`Styles`], says whether the
//!   shape's style is exclusive ([`Target`]: one `TSWP.ShapeInfoArchive`
//!   references it and no shape style names it as parent), and calls
//!   [`plan_set`] / [`plan_reset`] with [`Slot`] writes (a
//!   [`ShapeProperty`] gives them). It publishes the [`Plan`]:
//!   - [`Plan::Rewrite`]: replace the style's payload;
//!   - [`Plan::Derive`]: allocate an identifier; store the variation
//!     (message type [`MESSAGE_TYPE`], version [`MESSAGE_VERSION`], object
//!     references [`Variation::references`]) at the end of the stylesheet's
//!     component; repoint the shape ([`repoint`], and its ShapeInfo metadata
//!     reference); register it ([`stylesheet_with_variation`], and the
//!     stylesheet metadata reference); bind a UUID in the stylesheet
//!     component; add an external reference from the shape's component when
//!     it differs; raise the watermark to the new identifier;
//!   - [`Plan::Collapse`]: repoint the shape to the parent; unregister
//!     ([`stylesheet_without_variation`]) and remove the variation; release
//!     its UUID and every external reference to it; add an external
//!     reference to the parent from the shape's component when it differs;
//!     release the watermark suffix.
//! - **Workspace**: [`set`], [`reset`] and [`value`] do all of that over a
//!   [`Workspace`] ([`target`] and [`apply`] are
//!   the two halves).
//!
//! [`property`] holds [`property::Fills`], [`property::Strokes`],
//! [`property::Shadows`], [`property::ShapeEffects`],
//! [`property::TextLayout`] and [`property::TextColumns`]. Image fills, which
//! also move data references and media, and line endpoints, which the host
//! writes as a pair with the line geometry, use the plan layer with the
//! families' own values.

use std::collections::{BTreeMap, HashSet};
use std::mem::size_of;

use litchi_iwa_common::wire;
use litchi_iwa_core::{ArchiveObject, RawMessage};
use litchi_iwa_protos::chart_style_codec::Limits as CodecLimits;
use litchi_iwa_protos::shape_style_codec;

use crate::graph::{
    self, Error, Location, Result, Workspace, push_message_field, push_varint, push_varint_field,
};

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const MAX_INHERITANCE_DEPTH: usize = 64;

const STYLE_SUPER: u32 = 1;
const STYLE_OVERRIDE_COUNT: u32 = 10;
const STYLE_PROPERTIES: u32 = 11;
const TSS_PARENT: u32 = 3;
const TSS_VARIATION: u32 = 4;
const TSS_STYLESHEET: u32 = 5;
const SHAPE_INFO_SUPER: u32 = 1;
const DRAWABLE_STYLE: u32 = 2;
const REFERENCE_IDENTIFIER: u32 = 1;
const STYLESHEET_STYLES: u32 = 1;
const STYLESHEET_PARENT_CHILDREN: u32 = 5;
const CHILDREN_PARENT: u32 = 1;
const CHILDREN: u32 = 2;
const HEAD_LINE_END: u32 = 6;
const TAIL_LINE_END: u32 = 7;
const SHAPE_PROPERTY_FIELDS: std::ops::RangeInclusive<u32> = 1..=7;
const TEXT_PROPERTY_FIELDS: std::ops::RangeInclusive<u32> = 1..=6;

/// Residual resources admitted by one stylesheet registration transition.
///
/// This is crate-private because format owners consume the shared graph
/// engine rather than exposing a second public stylesheet API. The helper
/// keeps the existing shape byte-rewrite path, but all nested scans,
/// fallible-vector capacity and candidate output are admitted first.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RegistrationLimits {
    pub(crate) max_input_bytes: usize,
    pub(crate) max_fields: usize,
    pub(crate) max_output_bytes: usize,
    pub(crate) max_work_bytes: usize,
    pub(crate) max_allocations: usize,
    pub(crate) max_scratch_bytes: usize,
}

/// Resource dimensions owned by the stylesheet registration seam.
///
/// `litchi_iwa_common::LimitKind` has no allocation or scratch dimensions;
/// keeping those dimensions private avoids coercing them into rewrite work
/// until the chart-font owner has a matching typed adapter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum RegistrationLimitKind {
    InputBytes,
    Fields,
    OutputBytes,
    Work,
    Allocations,
    ScratchBytes,
}

/// Typed failure from a bounded stylesheet registration transition.
#[derive(Debug)]
pub(crate) enum RegistrationError {
    Limit {
        kind: RegistrationLimitKind,
        observed: usize,
        maximum: usize,
    },
    Allocation {
        resource: &'static str,
        amount: usize,
    },
    Wire(litchi_iwa_common::Error),
    Graph(Error),
    Invalid(&'static str),
}

type RegistrationResult<T> = std::result::Result<T, RegistrationError>;

#[cfg(test)]
mod registration_tests {
    use super::{
        RegistrationError, RegistrationLimitKind, RegistrationLimits,
        stylesheet_with_variation_admitted, stylesheet_without_variation_admitted,
    };
    use crate::graph::push_message_field;

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        crate::graph::push_varint_field(&mut output, 1, identifier);
        output
    }

    fn stylesheet() -> Vec<u8> {
        let mut output = Vec::new();
        push_message_field(&mut output, 1, &reference(11));
        push_message_field(&mut output, 1, &reference(12));
        let mut children = Vec::new();
        push_message_field(&mut children, 1, &reference(10));
        push_message_field(&mut children, 2, &reference(11));
        push_message_field(&mut children, 2, &reference(12));
        crate::graph::push_varint_field(&mut children, 99, 8);
        push_message_field(&mut output, 5, &children);
        let mut second_parent = Vec::new();
        push_message_field(&mut second_parent, 1, &reference(20));
        push_message_field(&mut second_parent, 2, &reference(21));
        push_message_field(&mut output, 5, &second_parent);
        crate::graph::push_varint_field(&mut output, 99, 7);
        output
    }

    fn limits() -> RegistrationLimits {
        RegistrationLimits {
            max_input_bytes: 4096,
            max_fields: 4096,
            max_output_bytes: 4096,
            max_work_bytes: 32_768,
            max_allocations: 4096,
            max_scratch_bytes: 32_768,
        }
    }

    #[test]
    fn bounded_registration_preserves_unknowns_and_multiple_children() {
        let source = stylesheet();
        let (registered, report) = stylesheet_with_variation_admitted(&source, 10, 13, limits())
            .expect("bounded registration");
        assert!(registered.windows(2).any(|window| window == [0x98, 0x06]));
        assert!(report.fields > 5);
        assert!(report.work_bytes > 0);
        assert!(report.allocations > 1);
        assert!(report.scratch_bytes > 0);
        assert!(report.output_bytes > 0);
        let (restored, removal) =
            stylesheet_without_variation_admitted(&registered, 10, 13, limits())
                .expect("bounded unregistration");
        assert_eq!(restored, source);
        assert!(removal.allocations > 1);
    }

    #[test]
    fn bounded_registration_refuses_output_before_candidate_buffer() {
        let source = stylesheet();
        let mut limited = limits();
        limited.max_output_bytes = 0;
        let error = stylesheet_with_variation_admitted(&source, 10, 13, limited)
            .expect_err("zero output must refuse before a candidate is allocated");
        match error {
            RegistrationError::Limit {
                kind,
                observed,
                maximum,
            } => {
                assert_eq!(kind, RegistrationLimitKind::OutputBytes);
                assert!(observed > maximum);
            },
            other => panic!("unexpected output refusal: {other:?}"),
        }
    }

    #[test]
    fn bounded_registration_refuses_each_zero_dimension_before_rewrite() {
        fn zero_fields(limits: &mut RegistrationLimits) {
            limits.max_fields = 0;
        }
        fn zero_work(limits: &mut RegistrationLimits) {
            limits.max_work_bytes = 0;
        }
        fn zero_allocations(limits: &mut RegistrationLimits) {
            limits.max_allocations = 0;
        }
        fn zero_scratch(limits: &mut RegistrationLimits) {
            limits.max_scratch_bytes = 0;
        }
        let source = stylesheet();
        let cases: &[(RegistrationLimitKind, fn(&mut RegistrationLimits))] = &[
            (RegistrationLimitKind::Fields, zero_fields),
            (RegistrationLimitKind::Work, zero_work),
            (RegistrationLimitKind::Allocations, zero_allocations),
            (RegistrationLimitKind::ScratchBytes, zero_scratch),
        ];
        for &(expected, mutate) in cases {
            let mut limited = limits();
            mutate(&mut limited);
            let error = stylesheet_with_variation_admitted(&source, 10, 13, limited)
                .expect_err("zero resource must refuse before rewriting");
            match error {
                RegistrationError::Limit {
                    kind,
                    observed,
                    maximum,
                } => {
                    assert_eq!(kind, expected);
                    assert!(observed > maximum);
                },
                other => panic!("unexpected registration refusal: {other:?}"),
            }
        }
    }

    #[test]
    fn bounded_registration_reports_nonzero_insufficient_dimensions() {
        fn one_fields(limits: &mut RegistrationLimits) {
            limits.max_fields = 1;
        }
        fn one_work(limits: &mut RegistrationLimits) {
            limits.max_work_bytes = 1;
        }
        fn one_allocations(limits: &mut RegistrationLimits) {
            limits.max_allocations = 1;
        }
        fn one_scratch(limits: &mut RegistrationLimits) {
            limits.max_scratch_bytes = 1;
        }
        let source = stylesheet();
        let cases: &[(RegistrationLimitKind, fn(&mut RegistrationLimits))] = &[
            (RegistrationLimitKind::Fields, one_fields),
            (RegistrationLimitKind::Work, one_work),
            (RegistrationLimitKind::Allocations, one_allocations),
            (RegistrationLimitKind::ScratchBytes, one_scratch),
        ];
        for &(expected, mutate) in cases {
            let mut limited = limits();
            mutate(&mut limited);
            let error = stylesheet_with_variation_admitted(&source, 10, 13, limited)
                .expect_err("insufficient resource must refuse before rewriting");
            match error {
                RegistrationError::Limit {
                    kind,
                    observed,
                    maximum,
                } => {
                    assert_eq!(kind, expected);
                    assert!(observed > maximum);
                },
                other => panic!("unexpected registration refusal: {other:?}"),
            }
        }
    }
}

/// Bounded work consumed by one stylesheet registration transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RegistrationReport {
    pub(crate) input_bytes: usize,
    pub(crate) output_bytes: usize,
    pub(crate) fields: usize,
    pub(crate) work_bytes: usize,
    pub(crate) allocations: usize,
    pub(crate) scratch_bytes: usize,
}

impl RegistrationLimits {
    pub(crate) const PERMISSIVE: Self = Self {
        max_input_bytes: litchi_iwa_common::WireLimits::MAX_INPUT_BYTES,
        max_fields: litchi_iwa_common::WireLimits::MAX_FIELDS,
        max_output_bytes: litchi_iwa_common::WireLimits::MAX_OUTPUT_BYTES,
        max_work_bytes: litchi_iwa_common::WireLimits::MAX_REWRITE_WORK,
        max_allocations: litchi_iwa_common::WireLimits::MAX_FIELDS,
        max_scratch_bytes: litchi_iwa_common::WireLimits::MAX_INPUT_BYTES,
    };
}

#[derive(Debug, Clone, Copy)]
struct RegistrationScan {
    input_bytes: usize,
    output_bytes: usize,
    fields: usize,
    work_bytes: usize,
    allocations: usize,
    scratch_bytes: usize,
}

impl RegistrationScan {
    const fn empty() -> Self {
        Self {
            input_bytes: 0,
            output_bytes: 0,
            fields: 0,
            work_bytes: 0,
            allocations: 0,
            scratch_bytes: 0,
        }
    }
}

fn registration_limit(
    kind: RegistrationLimitKind,
    observed: usize,
    maximum: usize,
) -> RegistrationError {
    RegistrationError::Limit {
        kind,
        observed,
        maximum,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> RegistrationError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => match kind {
            litchi_iwa_common::LimitKind::InputBytes => {
                registration_limit(RegistrationLimitKind::InputBytes, observed, limit)
            },
            litchi_iwa_common::LimitKind::Fields => {
                registration_limit(RegistrationLimitKind::Fields, observed, limit)
            },
            litchi_iwa_common::LimitKind::OutputBytes => {
                registration_limit(RegistrationLimitKind::OutputBytes, observed, limit)
            },
            litchi_iwa_common::LimitKind::RewriteWork => {
                registration_limit(RegistrationLimitKind::Work, observed, limit)
            },
            other => RegistrationError::Wire(litchi_iwa_common::Error::LimitExceeded {
                kind: other,
                observed,
                limit,
            }),
        },
        litchi_iwa_common::Error::Allocation { resource, amount } => {
            RegistrationError::Allocation { resource, amount }
        },
        other => RegistrationError::Wire(other),
    }
}

impl From<litchi_iwa_common::Error> for RegistrationError {
    fn from(error: litchi_iwa_common::Error) -> Self {
        map_wire_error(error)
    }
}

impl From<Error> for RegistrationError {
    fn from(error: Error) -> Self {
        Self::Graph(error)
    }
}

fn registration_add(
    current: &mut usize,
    amount: usize,
    maximum: usize,
    kind: RegistrationLimitKind,
) -> RegistrationResult<()> {
    let observed = current
        .checked_add(amount)
        .ok_or_else(|| registration_limit(kind, usize::MAX, maximum))?;
    if observed > maximum {
        return Err(registration_limit(kind, observed, maximum));
    }
    *current = observed;
    Ok(())
}

fn registration_admit(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: RegistrationLimitKind,
) -> RegistrationResult<()> {
    let observed = current
        .checked_add(amount)
        .ok_or_else(|| registration_limit(kind, usize::MAX, maximum))?;
    if observed > maximum {
        return Err(registration_limit(kind, observed, maximum));
    }
    Ok(())
}

fn registration_remaining(
    current: usize,
    maximum: usize,
    kind: RegistrationLimitKind,
) -> RegistrationResult<usize> {
    maximum
        .checked_sub(current)
        .filter(|remaining| *remaining > 0)
        .ok_or_else(|| registration_limit(kind, current.saturating_add(1), maximum))
}

fn registration_wire_limits(
    limits: RegistrationLimits,
    scan: &RegistrationScan,
) -> RegistrationResult<litchi_iwa_common::WireLimits> {
    let input = registration_remaining(
        scan.input_bytes,
        limits.max_input_bytes,
        RegistrationLimitKind::InputBytes,
    )?;
    let fields = registration_remaining(
        scan.fields,
        limits.max_fields,
        RegistrationLimitKind::Fields,
    )?;
    let work = registration_remaining(
        scan.work_bytes,
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?
    .min(litchi_iwa_common::WireLimits::MAX_REWRITE_WORK);
    litchi_iwa_common::WireLimits::DEFAULT
        .with_input_bytes(input)
        .and_then(|value| value.with_fields(fields))
        .and_then(|value| value.with_rewrite_work(work))
        .map_err(map_wire_error)
}

fn registration_fields(
    data: &[u8],
    limits: RegistrationLimits,
    scan: &mut RegistrationScan,
) -> RegistrationResult<Vec<wire::WireField>> {
    registration_admit(
        scan.input_bytes,
        data.len(),
        limits.max_input_bytes,
        RegistrationLimitKind::InputBytes,
    )?;
    registration_admit(
        scan.work_bytes,
        data.len(),
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    let wire_limits = registration_wire_limits(limits, scan)?;
    registration_add(
        &mut scan.scratch_bytes,
        size_of::<wire::WireField>(),
        limits.max_scratch_bytes,
        RegistrationLimitKind::ScratchBytes,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    let mut field_count = 0usize;
    wire::visit_wire_fields(data, wire_limits, |_field| {
        field_count = field_count.saturating_add(1);
        Ok(())
    })
    .map_err(map_wire_error)?;
    let second_work = data.len().checked_add(field_count).ok_or_else(|| {
        registration_limit(
            RegistrationLimitKind::Work,
            usize::MAX,
            limits.max_work_bytes,
        )
    })?;
    let first_work = second_work;
    registration_add(
        &mut scan.input_bytes,
        data.len(),
        limits.max_input_bytes,
        RegistrationLimitKind::InputBytes,
    )?;
    registration_add(
        &mut scan.fields,
        field_count,
        limits.max_fields,
        RegistrationLimitKind::Fields,
    )?;
    registration_add(
        &mut scan.work_bytes,
        first_work,
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    registration_admit(
        scan.input_bytes,
        data.len(),
        limits.max_input_bytes,
        RegistrationLimitKind::InputBytes,
    )?;
    registration_admit(
        scan.fields,
        field_count,
        limits.max_fields,
        RegistrationLimitKind::Fields,
    )?;
    registration_admit(
        scan.work_bytes,
        second_work,
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    let second_wire_limits = registration_wire_limits(limits, scan)?;
    let field_storage = field_count
        .checked_mul(size_of::<wire::WireField>())
        .ok_or_else(|| {
            registration_limit(
                RegistrationLimitKind::ScratchBytes,
                usize::MAX,
                limits.max_scratch_bytes,
            )
        })?;
    registration_add(
        &mut scan.scratch_bytes,
        field_storage,
        limits.max_scratch_bytes,
        RegistrationLimitKind::ScratchBytes,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    let fields = wire::parse_wire_fields(data, second_wire_limits).map_err(map_wire_error)?;
    registration_add(
        &mut scan.input_bytes,
        data.len(),
        limits.max_input_bytes,
        RegistrationLimitKind::InputBytes,
    )?;
    registration_add(
        &mut scan.fields,
        field_count,
        limits.max_fields,
        RegistrationLimitKind::Fields,
    )?;
    registration_add(
        &mut scan.work_bytes,
        second_work,
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    Ok(fields)
}

fn scan_reference(
    data: &[u8],
    limits: RegistrationLimits,
    scan: &mut RegistrationScan,
) -> RegistrationResult<()> {
    let fields = registration_fields(data, limits, scan)?;
    if fields
        .iter()
        .any(|field| field.number() == REFERENCE_IDENTIFIER && field.wire_type() != 0)
    {
        return Err(RegistrationError::Invalid(
            "a stylesheet reference has the wrong wire type",
        ));
    }
    Ok(())
}

fn scan_stylesheet(
    stylesheet: &[u8],
    limits: RegistrationLimits,
) -> RegistrationResult<RegistrationScan> {
    let mut scan = RegistrationScan::empty();
    let fields = registration_fields(stylesheet, limits, &mut scan)?;
    for field in fields {
        if field.wire_type() != 2
            || !matches!(
                field.number(),
                STYLESHEET_STYLES | STYLESHEET_PARENT_CHILDREN
            )
        {
            continue;
        }
        let payload = field.payload(stylesheet).map_err(map_wire_error)?;
        if field.number() == STYLESHEET_STYLES {
            scan_reference(payload, limits, &mut scan)?;
        } else {
            let children = registration_fields(payload, limits, &mut scan)?;
            for child in children {
                if child.number() == CHILDREN && child.wire_type() == 2 {
                    scan_reference(
                        child.payload(payload).map_err(map_wire_error)?,
                        limits,
                        &mut scan,
                    )?;
                }
            }
        }
    }
    Ok(scan)
}

fn bounded_copy(
    bytes: &[u8],
    scan: &mut RegistrationScan,
    limits: RegistrationLimits,
    resource: &'static str,
) -> RegistrationResult<Vec<u8>> {
    registration_add(
        &mut scan.scratch_bytes,
        bytes.len(),
        limits.max_scratch_bytes,
        RegistrationLimitKind::ScratchBytes,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|_| RegistrationError::Allocation {
            resource,
            amount: bytes.len(),
        })?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn registration_reserve<T>(
    output: &mut Vec<T>,
    additional: usize,
    scan: &mut RegistrationScan,
    limits: RegistrationLimits,
    resource: &'static str,
) -> RegistrationResult<()> {
    if additional == 0 {
        return Ok(());
    }
    let bytes = additional.checked_mul(size_of::<T>()).ok_or_else(|| {
        registration_limit(
            RegistrationLimitKind::ScratchBytes,
            usize::MAX,
            limits.max_scratch_bytes,
        )
    })?;
    registration_add(
        &mut scan.scratch_bytes,
        bytes,
        limits.max_scratch_bytes,
        RegistrationLimitKind::ScratchBytes,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    output
        .try_reserve_exact(additional)
        .map_err(|_| RegistrationError::Allocation {
            resource,
            amount: additional,
        })?;
    Ok(())
}

fn bounded_reference_identifier(
    data: &[u8],
    limits: RegistrationLimits,
    scan: &mut RegistrationScan,
) -> RegistrationResult<u64> {
    let fields = registration_fields(data, limits, scan)?;
    let field = fields
        .iter()
        .find(|field| field.number() == REFERENCE_IDENTIFIER)
        .ok_or(RegistrationError::Invalid(
            "a stylesheet reference has no identifier",
        ))?;
    if field.wire_type() != 0 {
        return Err(RegistrationError::Invalid(
            "a stylesheet reference has the wrong wire type",
        ));
    }
    let (identifier, _) = litchi_iwa_common::decode_varint_from_bytes(
        &data[field.key_end()..field.end()],
    )
    .map_err(|_| RegistrationError::Invalid("a stylesheet reference identifier is malformed"))?;
    Ok(identifier)
}

fn bounded_repeated_payloads(
    data: &[u8],
    number: u32,
    limits: RegistrationLimits,
    scan: &mut RegistrationScan,
) -> RegistrationResult<Vec<Vec<u8>>> {
    let fields = registration_fields(data, limits, scan)?;
    let count = fields
        .iter()
        .filter(|field| field.number() == number)
        .count();
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    let mut payloads = Vec::new();
    payloads
        .try_reserve_exact(count)
        .map_err(|_| RegistrationError::Allocation {
            resource: "stylesheet repeated payloads",
            amount: count,
        })?;
    for field in fields {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(RegistrationError::Invalid(
                "a repeated stylesheet field has the wrong wire type",
            ));
        }
        payloads.push(bounded_copy(
            field.payload(data).map_err(map_wire_error)?,
            scan,
            limits,
            "stylesheet repeated payload",
        )?);
    }
    Ok(payloads)
}

fn encoded_message_len(number: u32, payload_len: usize) -> RegistrationResult<usize> {
    let key = (u64::from(number) << 3) | 2;
    let length = u64::try_from(payload_len)
        .map_err(|_| RegistrationError::Invalid("stylesheet payload is too large"))?;
    litchi_iwa_common::varint::encoded_len(key)
        .checked_add(litchi_iwa_common::varint::encoded_len(length))
        .and_then(|value| value.checked_add(payload_len))
        .ok_or(RegistrationError::Invalid(
            "stylesheet output length overflow",
        ))
}

fn prepare_output(
    output_len: usize,
    scan: &mut RegistrationScan,
    limits: RegistrationLimits,
) -> RegistrationResult<Vec<u8>> {
    registration_add(
        &mut scan.output_bytes,
        output_len,
        limits.max_output_bytes,
        RegistrationLimitKind::OutputBytes,
    )?;
    registration_add(
        &mut scan.work_bytes,
        output_len,
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|_| RegistrationError::Allocation {
            resource: "stylesheet candidate output",
            amount: output_len,
        })?;
    Ok(output)
}

fn bounded_rewrite_repeated(
    data: &[u8],
    number: u32,
    replacements: &[Vec<u8>],
    limits: RegistrationLimits,
    scan: &mut RegistrationScan,
) -> RegistrationResult<Vec<u8>> {
    let fields = registration_fields(data, limits, scan)?;
    let matches = fields
        .iter()
        .filter(|field| field.number() == number)
        .count();
    for field in fields.iter().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(RegistrationError::Invalid(
                "a repeated stylesheet field has the wrong wire type",
            ));
        }
    }
    let removed = fields
        .iter()
        .filter(|field| field.number() == number)
        .try_fold(0usize, |total, field| {
            total.checked_add(field.end() - field.start())
        })
        .ok_or(RegistrationError::Invalid(
            "stylesheet output length overflow",
        ))?;
    let mut replacement_len = 0usize;
    for payload in replacements {
        replacement_len = replacement_len
            .checked_add(encoded_message_len(number, payload.len())?)
            .ok_or(RegistrationError::Invalid(
                "stylesheet output length overflow",
            ))?;
    }
    let output_len = data
        .len()
        .checked_sub(removed)
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or(RegistrationError::Invalid(
            "stylesheet output length overflow",
        ))?;
    let mut output = prepare_output(output_len, scan, limits)?;
    let mut replacement_index = 0usize;
    for field in fields {
        if field.number() == number {
            if let Some(payload) = replacements.get(replacement_index) {
                output.extend_from_slice(field.key(data).map_err(map_wire_error)?);
                push_varint(&mut output, payload.len() as u64);
                output.extend_from_slice(payload);
            }
            replacement_index += 1;
        } else {
            output.extend_from_slice(field.raw(data).map_err(map_wire_error)?);
        }
    }
    if matches == 0 {
        for payload in replacements {
            push_message_field(&mut output, number, payload);
        }
    } else if replacement_index < replacements.len() {
        for payload in &replacements[replacement_index..] {
            push_message_field(&mut output, number, payload);
        }
    }
    debug_assert_eq!(output.len(), output_len);
    Ok(output)
}

fn map_char_style_error(error: litchi_iwa_text_wire::CharStyleError) -> RegistrationError {
    match error {
        litchi_iwa_text_wire::CharStyleError::Limit {
            kind,
            observed,
            maximum,
        } => match kind {
            litchi_iwa_common::LimitKind::InputBytes => {
                registration_limit(RegistrationLimitKind::InputBytes, observed, maximum)
            },
            litchi_iwa_common::LimitKind::Fields => {
                registration_limit(RegistrationLimitKind::Fields, observed, maximum)
            },
            litchi_iwa_common::LimitKind::OutputBytes => {
                registration_limit(RegistrationLimitKind::OutputBytes, observed, maximum)
            },
            litchi_iwa_common::LimitKind::RewriteWork => {
                registration_limit(RegistrationLimitKind::Work, observed, maximum)
            },
            other => RegistrationError::Wire(litchi_iwa_common::Error::LimitExceeded {
                kind: other,
                observed,
                limit: maximum,
            }),
        },
        litchi_iwa_text_wire::CharStyleError::Allocation { amount } => {
            RegistrationError::Allocation {
                resource: "stylesheet text-wire output",
                amount,
            }
        },
        litchi_iwa_text_wire::CharStyleError::Invalid
        | litchi_iwa_text_wire::CharStyleError::Span => {
            RegistrationError::Invalid("invalid stylesheet text-wire payload")
        },
        _ => RegistrationError::Invalid("unsupported stylesheet text-wire failure"),
    }
}

fn bounded_register(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> RegistrationResult<(Vec<u8>, RegistrationReport)> {
    let mut scan = scan_stylesheet(stylesheet, limits)?;
    let styles = bounded_repeated_payloads(stylesheet, STYLESHEET_STYLES, limits, &mut scan)?;
    for payload in &styles {
        if bounded_reference_identifier(payload, limits, &mut scan)? == style {
            return Err(RegistrationError::Invalid(
                "the stylesheet already lists a new style",
            ));
        }
    }
    let parent_entries =
        bounded_repeated_payloads(stylesheet, STYLESHEET_PARENT_CHILDREN, limits, &mut scan)?;
    let mut parent_matches = 0usize;
    for entry in &parent_entries {
        let entry_parent = bounded_repeated_payloads(entry, CHILDREN_PARENT, limits, &mut scan)?
            .first()
            .map(|payload| bounded_reference_identifier(payload, limits, &mut scan))
            .transpose()?
            .unwrap_or(0);
        if entry_parent == parent {
            parent_matches += 1;
        }
    }
    if parent_matches > 1 {
        return Err(RegistrationError::Invalid(
            "the stylesheet repeats a parent style",
        ));
    }
    // The shared text-wire writer preserves insertion-after-last-style and
    // unknown/order bytes exactly. The preflight above validates ownership;
    // this second scan admits the writer's nested traversal against residual
    // input/field/work allowances before it allocates its candidate buffer.
    scan_stylesheet(stylesheet, limits).map(|next| scan = next)?;
    let wire_limits = registration_wire_limits(limits, &scan)?;
    let output =
        litchi_iwa_text_wire::stylesheet_with_style(stylesheet, style, Some(parent), wire_limits)
            .map_err(map_char_style_error)?;
    registration_add(
        &mut scan.output_bytes,
        output.len(),
        limits.max_output_bytes,
        RegistrationLimitKind::OutputBytes,
    )?;
    registration_add(
        &mut scan.work_bytes,
        output.len(),
        limits.max_work_bytes,
        RegistrationLimitKind::Work,
    )?;
    registration_add(
        &mut scan.allocations,
        1,
        limits.max_allocations,
        RegistrationLimitKind::Allocations,
    )?;
    Ok((
        output,
        RegistrationReport {
            input_bytes: scan.input_bytes,
            output_bytes: scan.output_bytes,
            fields: scan.fields,
            work_bytes: scan.work_bytes,
            allocations: scan.allocations,
            scratch_bytes: scan.scratch_bytes,
        },
    ))
}

fn bounded_unregister(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> RegistrationResult<(Vec<u8>, RegistrationReport)> {
    let mut scan = scan_stylesheet(stylesheet, limits)?;
    let styles = bounded_repeated_payloads(stylesheet, STYLESHEET_STYLES, limits, &mut scan)?;
    let mut listed = 0usize;
    let mut kept_styles = Vec::new();
    registration_reserve(
        &mut kept_styles,
        styles.len(),
        &mut scan,
        limits,
        "stylesheet kept styles",
    )?;
    for payload in styles {
        if bounded_reference_identifier(&payload, limits, &mut scan)? == style {
            listed += 1;
        } else {
            kept_styles.push(payload);
        }
    }
    if listed != 1 {
        return Err(RegistrationError::Invalid(
            "the stylesheet must list a removed style once",
        ));
    }
    let data = bounded_rewrite_repeated(
        stylesheet,
        STYLESHEET_STYLES,
        &kept_styles,
        limits,
        &mut scan,
    )?;
    let entries = bounded_repeated_payloads(&data, STYLESHEET_PARENT_CHILDREN, limits, &mut scan)?;
    let mut removed = 0usize;
    let mut kept_entries = Vec::new();
    registration_reserve(
        &mut kept_entries,
        entries.len(),
        &mut scan,
        limits,
        "stylesheet kept parent entries",
    )?;
    for entry in entries {
        let entry_parent = bounded_repeated_payloads(&entry, CHILDREN_PARENT, limits, &mut scan)?
            .first()
            .map(|payload| bounded_reference_identifier(payload, limits, &mut scan))
            .transpose()?
            .unwrap_or(0);
        let children = bounded_repeated_payloads(&entry, CHILDREN, limits, &mut scan)?;
        let mut kept_children = Vec::new();
        registration_reserve(
            &mut kept_children,
            children.len(),
            &mut scan,
            limits,
            "stylesheet kept children",
        )?;
        let mut found = false;
        for child in children {
            if bounded_reference_identifier(&child, limits, &mut scan)? == style {
                found = true;
                removed += 1;
            } else {
                kept_children.push(child);
            }
        }
        if found && entry_parent != parent {
            return Err(RegistrationError::Invalid(
                "the stylesheet maps a removed style under another parent",
            ));
        }
        if found && kept_children.is_empty() {
            continue;
        }
        kept_entries.push(if found {
            bounded_rewrite_repeated(&entry, CHILDREN, &kept_children, limits, &mut scan)?
        } else {
            entry
        });
    }
    if removed != 1 {
        return Err(RegistrationError::Invalid(
            "the stylesheet must map a removed style once",
        ));
    }
    let output = bounded_rewrite_repeated(
        &data,
        STYLESHEET_PARENT_CHILDREN,
        &kept_entries,
        limits,
        &mut scan,
    )?;
    Ok((
        output,
        RegistrationReport {
            input_bytes: scan.input_bytes,
            output_bytes: scan.output_bytes,
            fields: scan.fields,
            work_bytes: scan.work_bytes,
            allocations: scan.allocations,
            scratch_bytes: scan.scratch_bytes,
        },
    ))
}

/// Bounded stylesheet registration used by shape and chart-font owners.
pub(crate) fn stylesheet_with_variation_admitted(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> RegistrationResult<(Vec<u8>, RegistrationReport)> {
    bounded_register(stylesheet, parent, style, limits)
}

pub(crate) fn stylesheet_without_variation_admitted(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> RegistrationResult<(Vec<u8>, RegistrationReport)> {
    bounded_unregister(stylesheet, parent, style, limits)
}

fn registration_error_to_graph(error: RegistrationError) -> Error {
    match error {
        RegistrationError::Limit {
            kind: RegistrationLimitKind::InputBytes,
            observed,
            maximum,
        } => Error::Wire(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::InputBytes,
            observed,
            limit: maximum,
        }),
        RegistrationError::Limit {
            kind: RegistrationLimitKind::Fields,
            observed,
            maximum,
        } => Error::Wire(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::Fields,
            observed,
            limit: maximum,
        }),
        RegistrationError::Limit {
            kind: RegistrationLimitKind::OutputBytes,
            observed,
            maximum,
        } => Error::Wire(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::OutputBytes,
            observed,
            limit: maximum,
        }),
        RegistrationError::Limit {
            kind: RegistrationLimitKind::Work,
            observed,
            maximum,
        } => Error::Wire(litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::RewriteWork,
            observed,
            limit: maximum,
        }),
        RegistrationError::Limit {
            kind: RegistrationLimitKind::Allocations,
            observed,
            maximum,
        } => Error::Limit {
            kind: graph::GraphLimitKind::AllocationEvents,
            observed,
            maximum,
        },
        RegistrationError::Limit {
            kind: RegistrationLimitKind::ScratchBytes,
            observed,
            maximum,
        } => Error::Limit {
            kind: graph::GraphLimitKind::ScratchBytes,
            observed,
            maximum,
        },
        RegistrationError::Allocation { resource, amount } => {
            Error::Wire(litchi_iwa_common::Error::Allocation { resource, amount })
        },
        RegistrationError::Wire(error) => Error::Wire(error),
        RegistrationError::Graph(error) => error,
        RegistrationError::Invalid(reason) => Error::Invalid(reason),
    }
}

/// Compatibility adapter for existing shape callers. New bounded owners
/// should consume [`stylesheet_with_variation_admitted`] to retain allocation
/// and scratch dimensions instead of this legacy graph-error projection.
pub(crate) fn stylesheet_with_variation_bounded(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> Result<(Vec<u8>, RegistrationReport)> {
    stylesheet_with_variation_admitted(stylesheet, parent, style, limits)
        .map_err(registration_error_to_graph)
}

pub(crate) fn stylesheet_without_variation_bounded(
    stylesheet: &[u8],
    parent: u64,
    style: u64,
    limits: RegistrationLimits,
) -> Result<(Vec<u8>, RegistrationReport)> {
    stylesheet_without_variation_admitted(stylesheet, parent, style, limits)
        .map_err(registration_error_to_graph)
}

/// One drawing property of `TSD.ShapeStylePropertiesArchive`.
pub trait ShapeProperty {
    /// The semantic value.
    type Value: Clone + PartialEq;

    /// The slot groups the value is stored in. Each group is inherited as a
    /// unit from the first style on the parent chain that defines any of its
    /// slots (a presence flag and its payload, such as `padding_null` and
    /// `padding`, form one group); a reset clears every slot.
    fn groups(&self) -> &'static [&'static [Slot]];

    /// The value of the stored slots, in the flattened order of
    /// [`Self::groups`] (`None`: no style on the chain defines the slot's
    /// group, or the defining style leaves the slot out).
    ///
    /// # Errors
    ///
    /// A stored value is invalid or unsupported.
    fn read(&self, stored: &[Option<&[u8]>]) -> Result<Self::Value>;

    /// The slot writes that store `value` (`None` clears a slot).
    ///
    /// # Errors
    ///
    /// The value cannot be stored by this engine.
    fn writes(&self, value: &Self::Value) -> Result<Vec<(Slot, Option<Value>)>>;

    /// Whether a reset of a style that is not a minimal variation reports
    /// "nothing to reset" (`true`) or fails (`false`).
    fn lenient_reset(&self) -> bool {
        false
    }

    /// Refuse to replace or remove a stored slot this engine cannot release
    /// (an image fill's data reference).
    ///
    /// # Errors
    ///
    /// The stored payload is not replaceable here.
    fn replaceable(&self, slot: Slot, payload: &[u8]) -> Result<()> {
        let _ = (slot, payload);
        Ok(())
    }
}

/// One raw override: wire type and value bytes (a length-delimited payload,
/// a varint's bytes or a fixed32's four bytes).
type Raw = (u8, Vec<u8>);

/// The direct overrides of a minimal variation.
#[derive(Debug, Clone, Default)]
struct Overrides {
    shape: BTreeMap<u32, Raw>,
    text: BTreeMap<u32, Raw>,
}

impl Overrides {
    fn count(&self) -> u64 {
        (self.shape.len() + self.text.len()) as u64
    }

    fn is_empty(&self) -> bool {
        self.count() == 0
    }

    fn get(&self, slot: Slot) -> Option<&Raw> {
        let (text, number) = slot.location();
        if text {
            self.text.get(&number)
        } else {
            self.shape.get(&number)
        }
    }

    fn set(&mut self, slot: Slot, value: Option<&Value>) {
        let (text, number) = slot.location();
        let map = if text {
            &mut self.text
        } else {
            &mut self.shape
        };
        match value {
            Some(value) => {
                map.insert(number, value.0.clone());
            },
            None => {
                map.remove(&number);
            },
        }
    }
}

fn push_raw(output: &mut Vec<u8>, number: u32, (wire_type, bytes): &Raw) {
    push_varint(output, (u64::from(number) << 3) | u64::from(*wire_type));
    if *wire_type == 2 {
        push_varint(output, bytes.len() as u64);
    }
    output.extend_from_slice(bytes);
}

fn raw_of(data: &[u8], field: wire::WireField) -> Result<Raw> {
    let start = if field.wire_type() == 2 {
        field.payload_start()
    } else {
        field.key_end()
    };
    Ok((
        field.wire_type(),
        data.get(start..field.end())
            .ok_or(Error::Invalid("a style field is truncated"))?
            .to_vec(),
    ))
}

fn single(data: &[u8], number: u32) -> Result<Option<&[u8]>> {
    let payloads = graph::repeated(data, number)?;
    match payloads.as_slice() {
        [] => Ok(None),
        [payload] => Ok(Some(payload)),
        _ => Err(Error::Invalid("a singular style field is repeated")),
    }
}

fn reference_identifier(reference: &[u8]) -> Result<u64> {
    graph::varint(reference, REFERENCE_IDENTIFIER)?
        .ok_or(Error::Invalid("a reference has no identifier"))
}

fn field_numbers(data: &[u8]) -> Result<Vec<u32>> {
    let mut numbers = graph::fields(data)?
        .into_iter()
        .map(wire::WireField::number)
        .collect::<Vec<_>>();
    numbers.sort_unstable();
    Ok(numbers)
}

/// The parts of a `TSWP.ShapeStyleArchive` payload the engine reads.
struct Style<'a> {
    parent: Option<u64>,
    stylesheet: Option<u64>,
    shape_properties: Option<&'a [u8]>,
}

fn style_parts(payload: &[u8]) -> Result<Style<'_>> {
    let tsd = single(payload, STYLE_SUPER)?
        .ok_or(Error::Invalid("a shape style has no drawing style"))?;
    let tss =
        single(tsd, STYLE_SUPER)?.ok_or(Error::Invalid("a shape style has no style parent"))?;
    Ok(Style {
        parent: single(tss, TSS_PARENT)?
            .map(reference_identifier)
            .transpose()?,
        stylesheet: single(tss, TSS_STYLESHEET)?
            .map(reference_identifier)
            .transpose()?,
        shape_properties: single(tsd, STYLE_PROPERTIES)?,
    })
}

/// The direct overrides when the style is a minimal variation (the host's
/// `direct_shape_style_overrides`).
fn direct_overrides(payload: &[u8]) -> Result<Option<Overrides>> {
    if field_numbers(payload)? != [STYLE_SUPER, STYLE_OVERRIDE_COUNT, STYLE_PROPERTIES] {
        return Ok(None);
    }
    let tsd = single(payload, STYLE_SUPER)?
        .ok_or(Error::Invalid("a shape style has no drawing style"))?;
    if field_numbers(tsd)? != [STYLE_SUPER, STYLE_OVERRIDE_COUNT, STYLE_PROPERTIES] {
        return Ok(None);
    }
    let tss =
        single(tsd, STYLE_SUPER)?.ok_or(Error::Invalid("a shape style has no style parent"))?;
    if field_numbers(tss)? != [TSS_PARENT, TSS_VARIATION, TSS_STYLESHEET]
        || graph::varint(tss, TSS_VARIATION)? != Some(1)
    {
        return Ok(None);
    }
    let mut overrides = Overrides::default();
    let shape = single(tsd, STYLE_PROPERTIES)?.unwrap_or_default();
    for field in graph::fields(shape)? {
        if !SHAPE_PROPERTY_FIELDS.contains(&field.number())
            || overrides
                .shape
                .insert(field.number(), raw_of(shape, field)?)
                .is_some()
        {
            return Ok(None);
        }
    }
    let text = single(payload, STYLE_PROPERTIES)?.unwrap_or_default();
    for field in graph::fields(text)? {
        if !TEXT_PROPERTY_FIELDS.contains(&field.number())
            || overrides
                .text
                .insert(field.number(), raw_of(text, field)?)
                .is_some()
        {
            return Ok(None);
        }
    }
    let count = overrides.count();
    let paired = overrides.shape.contains_key(&HEAD_LINE_END)
        == overrides.shape.contains_key(&TAIL_LINE_END);
    let semantic = count > 0
        && paired
        && graph::varint(payload, STYLE_OVERRIDE_COUNT)? == Some(count)
        && graph::varint(tsd, STYLE_OVERRIDE_COUNT)? == Some(count);
    Ok(semantic.then_some(overrides))
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut reference = Vec::with_capacity(11);
    push_varint_field(&mut reference, REFERENCE_IDENTIFIER, identifier);
    reference
}

/// The variation payload a protobuf encoder writes for `overrides` (the
/// host's `shape_style_variation_object`).
fn variation_payload(parent: u64, stylesheet: u64, overrides: &Overrides) -> Result<Vec<u8>> {
    let count = overrides.count();
    if count == 0 {
        return Err(Error::Invalid("a shape-style variation needs an override"));
    }
    let mut tss = Vec::new();
    push_message_field(&mut tss, TSS_PARENT, &reference(parent));
    push_varint_field(&mut tss, TSS_VARIATION, 1);
    push_message_field(&mut tss, TSS_STYLESHEET, &reference(stylesheet));
    let mut shape = Vec::new();
    for (number, raw) in &overrides.shape {
        push_raw(&mut shape, *number, raw);
    }
    let mut tsd = Vec::new();
    push_message_field(&mut tsd, STYLE_SUPER, &tss);
    push_varint_field(&mut tsd, STYLE_OVERRIDE_COUNT, count);
    push_message_field(&mut tsd, STYLE_PROPERTIES, &shape);
    let mut text = Vec::new();
    for (number, raw) in &overrides.text {
        push_raw(&mut text, *number, raw);
    }
    let mut payload = Vec::new();
    push_message_field(&mut payload, STYLE_SUPER, &tsd);
    push_varint_field(&mut payload, STYLE_OVERRIDE_COUNT, count);
    push_message_field(&mut payload, STYLE_PROPERTIES, &text);
    Ok(payload)
}

/// `TSWP.ShapeStyleArchive`.
pub const MESSAGE_TYPE: u32 = SHAPE_STYLE_MESSAGE_TYPE;
/// The message version a new variation is stored with.
pub const MESSAGE_VERSION: [u32; 3] = STANDARD_MESSAGE_VERSION;

/// One direct override slot of a shape-style variation: a field of
/// `TSD.ShapeStylePropertiesArchive` (drawing) or of
/// `TSWP.ShapeStylePropertiesArchive` (text layout).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Slot {
    /// `fill` (a `TSD.FillArchive`).
    Fill,
    /// `stroke` (a `TSD.StrokeArchive`).
    Stroke,
    /// `opacity` (a float).
    Opacity,
    /// `shadow` (a `TSD.ShadowArchive`).
    Shadow,
    /// `reflection` (a `TSD.ReflectionArchive`).
    Reflection,
    /// `head_line_end` (a `TSD.LineEndArchive`; written with the tail).
    HeadLineEnd,
    /// `tail_line_end` (a `TSD.LineEndArchive`; written with the head).
    TailLineEnd,
    /// `shrink_to_fit` (a bool).
    ShrinkToFit,
    /// `vertical_alignment` (an int32).
    VerticalAlignment,
    /// `columns_null` (a bool).
    ColumnsNull,
    /// `columns` (a `TSWP.ColumnsArchive`).
    Columns,
    /// `padding_null` (a bool).
    PaddingNull,
    /// `padding` (a `TSWP.PaddingArchive`).
    Padding,
}

impl Slot {
    /// `(text layout?, field number)`.
    const fn location(self) -> (bool, u32) {
        match self {
            Self::Fill => (false, 1),
            Self::Stroke => (false, 2),
            Self::Opacity => (false, 3),
            Self::Shadow => (false, 4),
            Self::Reflection => (false, 5),
            Self::HeadLineEnd => (false, HEAD_LINE_END),
            Self::TailLineEnd => (false, TAIL_LINE_END),
            Self::ShrinkToFit => (true, 1),
            Self::VerticalAlignment => (true, 2),
            Self::ColumnsNull => (true, 3),
            Self::Columns => (true, 4),
            Self::PaddingNull => (true, 5),
            Self::Padding => (true, 6),
        }
    }

    const ALL: [Self; 13] = [
        Self::Fill,
        Self::Stroke,
        Self::Opacity,
        Self::Shadow,
        Self::Reflection,
        Self::HeadLineEnd,
        Self::TailLineEnd,
        Self::ShrinkToFit,
        Self::VerticalAlignment,
        Self::ColumnsNull,
        Self::Columns,
        Self::PaddingNull,
        Self::Padding,
    ];
}

/// One stored override value: an encoded message, or a scalar.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Value(Raw);

impl Value {
    /// An encoded message (fill, stroke, shadow, reflection, line end,
    /// columns, padding).
    #[must_use]
    pub const fn message(encoded: Vec<u8>) -> Self {
        Self((2, encoded))
    }

    /// A float (opacity).
    #[must_use]
    pub fn float(value: f32) -> Self {
        Self((5, value.to_le_bytes().to_vec()))
    }

    /// A bool (`shrink_to_fit`, `columns_null`, `padding_null`).
    #[must_use]
    pub fn bool(value: bool) -> Self {
        Self((0, vec![u8::from(value)]))
    }

    /// An int32 (`vertical_alignment`).
    #[must_use]
    pub fn int32(value: i32) -> Self {
        let mut bytes = Vec::with_capacity(10);
        push_varint(&mut bytes, i64::from(value) as u64);
        Self((0, bytes))
    }

    /// The stored bytes: a message payload, a varint's bytes or a float's
    /// four bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.0.1
    }
}

/// Resolves shape-style payloads (`TSWP.ShapeStyleArchive`, message type
/// [`MESSAGE_TYPE`]) by object identifier.
pub trait Styles {
    /// The payload of style `identifier`, if it exists.
    fn shape_style(&self, identifier: u64) -> Option<&[u8]>;
}

/// The shape's style and whether only that shape uses it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Target {
    style: u64,
    exclusive: bool,
}

impl Target {
    /// `style` is the shape's style (see [`style_of`]); `exclusive` holds
    /// when exactly one `TSWP.ShapeInfoArchive` references it and no shape
    /// style names it as parent.
    #[must_use]
    pub const fn new(style: u64, exclusive: bool) -> Self {
        Self { style, exclusive }
    }

    /// The shape's style.
    #[must_use]
    pub const fn style(self) -> u64 {
        self.style
    }

    /// Whether only the shape uses the style.
    #[must_use]
    pub const fn exclusive(self) -> bool {
        self.exclusive
    }
}

/// A canonical shape-style variation payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Variation {
    payload: Vec<u8>,
    parent: u64,
    stylesheet: u64,
}

impl Variation {
    fn new(parent: u64, stylesheet: u64, overrides: &Overrides) -> Result<Self> {
        Ok(Self {
            payload: variation_payload(parent, stylesheet, overrides)?,
            parent,
            stylesheet,
        })
    }

    /// The `TSWP.ShapeStyleArchive` payload.
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Take the payload.
    #[must_use]
    pub fn into_payload(self) -> Vec<u8> {
        self.payload
    }

    /// The parent style.
    #[must_use]
    pub const fn parent(&self) -> u64 {
        self.parent
    }

    /// The stylesheet that registers the variation.
    #[must_use]
    pub const fn stylesheet(&self) -> u64 {
        self.stylesheet
    }

    /// The message's object references in archive-metadata order (the
    /// parent).
    #[must_use]
    pub fn references(&self) -> [u64; 1] {
        [self.parent]
    }
}

/// What a write does to the shape's style.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum Plan {
    /// Nothing to reset.
    Unchanged,
    /// Replace the payload of the shape's own variation in place (its
    /// identifier, parent, stylesheet and metadata stay).
    Rewrite(Variation),
    /// Store a new child variation of the shape's style (message type
    /// [`MESSAGE_TYPE`], version [`MESSAGE_VERSION`], object references
    /// [`Variation::references`]) in the stylesheet's component, repoint the
    /// shape to it ([`repoint`]) and register it
    /// ([`stylesheet_with_variation`]).
    Derive(Variation),
    /// Remove the shape's variation, which no longer overrides anything:
    /// repoint the shape to `parent` and unregister the variation from
    /// `stylesheet` ([`stylesheet_without_variation`]).
    Collapse {
        /// The variation's parent.
        parent: u64,
        /// The stylesheet that registers the variation.
        stylesheet: u64,
    },
}

impl Plan {
    /// Whether the plan changes anything.
    #[must_use]
    pub const fn changed(&self) -> bool {
        !matches!(self, Self::Unchanged)
    }
}

fn stored_slot(payload: &[u8], slot: Slot) -> Result<Option<&[u8]>> {
    let (text, number) = slot.location();
    let properties = if text {
        single(payload, STYLE_PROPERTIES)?
    } else {
        style_parts(payload)?.shape_properties
    };
    let Some(properties) = properties else {
        return Ok(None);
    };
    let mut found = None;
    for field in graph::fields(properties)? {
        if field.number() == number {
            if found.is_some() {
                return Err(Error::Invalid("a style property is repeated"));
            }
            let start = if field.wire_type() == 2 {
                field.payload_start()
            } else {
                field.key_end()
            };
            found = Some(
                properties
                    .get(start..field.end())
                    .ok_or(Error::Invalid("a style property is truncated"))?,
            );
        }
    }
    Ok(found)
}

fn styled(styles: &(impl Styles + ?Sized), style: u64) -> Result<&[u8]> {
    styles.shape_style(style).ok_or(Error::MissingObject(style))
}

/// The stored value of `slot` for a shape styled by `style`: the first
/// style on the parent chain that defines it (`None`: none does).
///
/// # Errors
///
/// A missing or malformed style, a cycle or a chain deeper than 64.
pub fn effective(
    styles: &(impl Styles + ?Sized),
    style: u64,
    slot: Slot,
) -> Result<Option<Vec<u8>>> {
    let mut visited = HashSet::new();
    let mut current = Some(style);
    for _ in 0..MAX_INHERITANCE_DEPTH {
        let Some(identifier) = current else {
            return Ok(None);
        };
        if !visited.insert(identifier) {
            return Err(Error::Invalid("a shape-style parent chain has a cycle"));
        }
        let payload = styled(styles, identifier)?;
        if let Some(value) = stored_slot(payload, slot)? {
            return Ok(Some(value.to_vec()));
        }
        current = style_parts(payload)?.parent;
    }
    Err(Error::Invalid("a shape-style parent chain is too deep"))
}

/// The slots a style payload overrides directly when it is a minimal
/// variation the engine may rewrite (`None` otherwise).
///
/// # Errors
///
/// A malformed payload.
pub fn direct_slots(style: &[u8]) -> Result<Option<Vec<Slot>>> {
    Ok(direct_overrides(style)?.map(|overrides| {
        Slot::ALL
            .into_iter()
            .filter(|slot| overrides.get(*slot).is_some())
            .collect()
    }))
}

/// Whether a style payload defines `slot` itself.
///
/// # Errors
///
/// A malformed payload.
pub fn defines(style: &[u8], slot: Slot) -> Result<bool> {
    Ok(stored_slot(style, slot)?.is_some())
}

/// The style a `TSWP.ShapeInfoArchive` payload references.
///
/// # Errors
///
/// A malformed payload or no style.
pub fn style_of(shape_info: &[u8]) -> Result<u64> {
    shape_style_codec::style_reference(shape_info, CodecLimits::DEFAULT)?
        .ok_or(Error::Invalid("a shape has no style"))
}

/// The parent of a shape-style payload.
///
/// # Errors
///
/// A malformed payload.
pub fn parent_of(style: &[u8]) -> Result<Option<u64>> {
    Ok(style_parts(style)?.parent)
}

/// The stylesheet of a shape-style payload.
///
/// # Errors
///
/// A malformed payload.
pub fn stylesheet_of(style: &[u8]) -> Result<Option<u64>> {
    Ok(style_parts(style)?.stylesheet)
}

/// Plan the writes `writes` (`Some`: set the slot; `None`: clear it) on the
/// shape styled by `target`: rewrite its exclusive minimal variation (or
/// collapse it when nothing is left), else derive a child variation with
/// the `Some` writes.
///
/// # Errors
///
/// A missing or malformed style, a style without stylesheet, a variation
/// without parent, or a derived variation with no override.
pub fn plan_set(
    styles: &(impl Styles + ?Sized),
    target: Target,
    writes: &[(Slot, Option<Value>)],
) -> Result<Plan> {
    let payload = styled(styles, target.style)?;
    let parts = style_parts(payload)?;
    let stylesheet = parts
        .stylesheet
        .ok_or(Error::Invalid("a shape style has no stylesheet"))?;
    if target.exclusive
        && let Some(mut overrides) = direct_overrides(payload)?
    {
        let parent = parts
            .parent
            .ok_or(Error::Invalid("a shape-style variation has no parent"))?;
        for (slot, value) in writes {
            overrides.set(*slot, value.as_ref());
        }
        if overrides.is_empty() {
            return Ok(Plan::Collapse { parent, stylesheet });
        }
        return Ok(Plan::Rewrite(Variation::new(
            parent, stylesheet, &overrides,
        )?));
    }
    let mut overrides = Overrides::default();
    for (slot, value) in writes {
        overrides.set(*slot, value.as_ref());
    }
    Ok(Plan::Derive(Variation::new(
        target.style,
        stylesheet,
        &overrides,
    )?))
}

/// Plan the reset of `slots` on the shape styled by `target`: unchanged
/// when its style defines none of them; else clear them from its exclusive
/// minimal variation (collapsing it when nothing is left), or derive a child
/// variation that stores `inherited` (the parent's effective values).
///
/// # Errors
///
/// As [`plan_set`], plus a style that defines a slot but is not a minimal
/// variation.
pub fn plan_reset(
    styles: &(impl Styles + ?Sized),
    target: Target,
    slots: &[Slot],
    inherited: &[(Slot, Value)],
) -> Result<Plan> {
    let payload = styled(styles, target.style)?;
    let mut defined = false;
    for slot in slots {
        defined |= stored_slot(payload, *slot)?.is_some();
    }
    if !defined {
        return Ok(Plan::Unchanged);
    }
    let parts = style_parts(payload)?;
    let stylesheet = parts
        .stylesheet
        .ok_or(Error::Invalid("a shape style has no stylesheet"))?;
    let mut direct = direct_overrides(payload)?.ok_or(Error::Invalid(
        "a shape style has the property plus unsupported direct overrides",
    ))?;
    let parent = parts
        .parent
        .ok_or(Error::Invalid("a shape-style variation has no parent"))?;
    if target.exclusive {
        for slot in slots {
            direct.set(*slot, None);
        }
        if direct.is_empty() {
            return Ok(Plan::Collapse { parent, stylesheet });
        }
        return Ok(Plan::Rewrite(Variation::new(parent, stylesheet, &direct)?));
    }
    let mut overrides = Overrides::default();
    for (slot, value) in inherited {
        overrides.set(*slot, Some(value));
    }
    Ok(Plan::Derive(Variation::new(
        target.style,
        stylesheet,
        &overrides,
    )?))
}

/// Repoint a `TSWP.ShapeInfoArchive` payload's style reference from `from`
/// to `to`, keeping every other byte (the archive metadata's object
/// reference moves with it).
///
/// # Errors
///
/// A malformed payload, or a style reference other than `from`.
pub fn repoint(shape_info: &[u8], from: u64, to: u64) -> Result<Vec<u8>> {
    wire::transform_length_delimited_field(shape_info, SHAPE_INFO_SUPER, |shape| {
        wire::transform_length_delimited_field(shape, DRAWABLE_STYLE, |style| -> Result<Vec<u8>> {
            if reference_identifier(style)? != from {
                return Err(Error::Invalid("a shape's style reference changed"));
            }
            Ok(wire::patch_varint_field(
                style,
                REFERENCE_IDENTIFIER,
                true,
                Some(to),
            )?)
        })
    })
}

/// Register one stylesheet variation under an explicit residual budget.
pub fn stylesheet_with_variation(stylesheet: &[u8], parent: u64, style: u64) -> Result<Vec<u8>> {
    stylesheet_with_variation_bounded(stylesheet, parent, style, RegistrationLimits::PERMISSIVE)
        .map(|(output, _report)| output)
}

pub fn stylesheet_without_variation(stylesheet: &[u8], parent: u64, style: u64) -> Result<Vec<u8>> {
    stylesheet_without_variation_bounded(stylesheet, parent, style, RegistrationLimits::PERMISSIVE)
        .map(|(output, _report)| output)
}

fn shape_style_reference(workspace: &Workspace<'_>, drawable: u64) -> Result<(usize, u64)> {
    let (message, payload) = workspace.unique_payload(
        drawable,
        SHAPE_INFO_MESSAGE_TYPE,
        "a shape must have exactly one ShapeInfo payload",
    )?;
    let style = shape_style_codec::style_reference(payload, CodecLimits::DEFAULT)?
        .ok_or(Error::Invalid("a shape has no style"))?;
    Ok((message, style))
}

fn style_payload<'w>(workspace: &'w Workspace<'_>, style: u64) -> Result<(usize, &'w [u8])> {
    workspace.unique_payload(
        style,
        SHAPE_STYLE_MESSAGE_TYPE,
        "a shape style must have exactly one ShapeStyle payload",
    )
}

fn inherited<P: ShapeProperty>(
    styles: &(impl Styles + ?Sized),
    first: u64,
    property: &P,
) -> Result<P::Value> {
    let mut values: Vec<Option<Vec<u8>>> = Vec::new();
    for group in property.groups() {
        values.extend(effective_group(styles, first, group)?);
    }
    let borrowed = values.iter().map(Option::as_deref).collect::<Vec<_>>();
    property.read(&borrowed)
}

/// The stored values of `group`'s slots from the first style on the parent
/// chain of `style` that defines any of them.
///
/// # Errors
///
/// A missing or malformed style, a cycle or a chain deeper than 64.
pub fn effective_group(
    styles: &(impl Styles + ?Sized),
    style: u64,
    group: &[Slot],
) -> Result<Vec<Option<Vec<u8>>>> {
    let mut visited = HashSet::new();
    let mut current = Some(style);
    for _ in 0..MAX_INHERITANCE_DEPTH {
        let Some(identifier) = current else {
            return Ok(vec![None; group.len()]);
        };
        if !visited.insert(identifier) {
            return Err(Error::Invalid("a shape-style parent chain has a cycle"));
        }
        let payload = styled(styles, identifier)?;
        let found = group
            .iter()
            .map(|slot| Ok(stored_slot(payload, *slot)?.map(<[u8]>::to_vec)))
            .collect::<Result<Vec<_>>>()?;
        if found.iter().any(Option::is_some) {
            return Ok(found);
        }
        current = style_parts(payload)?.parent;
    }
    Err(Error::Invalid("a shape-style parent chain is too deep"))
}

/// The effective value of `property` for shape drawable `drawable`.
///
/// # Errors
///
/// An invalid shape or style chain, or a value the property rejects.
pub fn value<P: ShapeProperty>(
    workspace: &Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<P::Value> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    inherited(workspace, style, property)
}

/// Whether exactly one shape uses `style` and no style inherits from it.
fn is_exclusive(workspace: &Workspace<'_>, style: u64) -> Result<bool> {
    let mut shapes = 0usize;
    for location in workspace.locations() {
        for object in &workspace.archive_at(location)?.objects {
            for message in &object.messages {
                match message.type_ {
                    SHAPE_INFO_MESSAGE_TYPE => {
                        if shape_style_codec::style_reference(&message.data, CodecLimits::DEFAULT)?
                            == Some(style)
                        {
                            shapes += 1;
                        }
                    },
                    SHAPE_STYLE_MESSAGE_TYPE => {
                        if style_parts(&message.data)?.parent == Some(style) {
                            return Ok(false);
                        }
                    },
                    _ => {},
                }
            }
        }
    }
    Ok(shapes == 1)
}

struct Context {
    drawable: u64,
    drawable_location: Location,
    style: u64,
    style_location: Location,
    stylesheet: u64,
}

fn context(workspace: &Workspace<'_>, drawable: u64, style: u64) -> Result<Context> {
    let (_, payload) = style_payload(workspace, style)?;
    let stylesheet = style_parts(payload)?
        .stylesheet
        .ok_or(Error::Invalid("a shape style has no stylesheet"))?;
    let style_location = workspace.locate(style)?;
    if workspace.locate(stylesheet)? != style_location {
        return Err(Error::Invalid(
            "a shape style is not stored with its stylesheet",
        ));
    }
    Ok(Context {
        drawable,
        drawable_location: workspace.locate(drawable)?,
        style,
        style_location,
        stylesheet,
    })
}

fn repoint_drawable(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    from: u64,
    to: u64,
) -> Result<()> {
    let (message, payload) = workspace.unique_payload(
        drawable,
        SHAPE_INFO_MESSAGE_TYPE,
        "a shape must have exactly one ShapeInfo payload",
    )?;
    let data = repoint(payload, from, to)?;
    workspace.replace_payload(drawable, message, data)?;
    let info = workspace
        .object_mut(drawable)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a ShapeInfo payload has no metadata"))?;
    let mut replaced = 0usize;
    for identifier in &mut info.object_references {
        if *identifier == from {
            *identifier = to;
            replaced += 1;
        }
    }
    for field in &mut info.field_infos {
        for identifier in &mut field.object_references {
            if *identifier == from {
                *identifier = to;
            }
        }
    }
    if replaced != 1 {
        return Err(Error::Invalid(
            "a shape's metadata must reference its style once",
        ));
    }
    Ok(())
}

fn register(workspace: &mut Workspace<'_>, stylesheet: u64, parent: u64, style: u64) -> Result<()> {
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let data = stylesheet_with_variation(payload, parent, style)?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    if info.object_references.contains(&style) {
        return Err(Error::Invalid(
            "the stylesheet metadata already references a new style",
        ));
    }
    info.object_references.push(style);
    Ok(())
}

fn unregister(
    workspace: &mut Workspace<'_>,
    stylesheet: u64,
    parent: u64,
    style: u64,
) -> Result<()> {
    let (message, payload) = workspace.unique_payload(
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        "a stylesheet must have exactly one payload",
    )?;
    let data = stylesheet_without_variation(payload, parent, style)?;
    workspace.replace_payload(stylesheet, message, data)?;
    let info = workspace
        .object_mut(stylesheet)?
        .archive_info
        .message_infos
        .get_mut(message)
        .ok_or(Error::Invalid("a stylesheet payload has no metadata"))?;
    if info
        .object_references
        .iter()
        .filter(|id| **id == style)
        .count()
        != 1
    {
        return Err(Error::Invalid(
            "the stylesheet metadata must reference a removed style once",
        ));
    }
    info.object_references.retain(|id| *id != style);
    for field in &mut info.field_infos {
        field.object_references.retain(|id| *id != style);
    }
    Ok(())
}

fn rewrite_variation(
    workspace: &mut Workspace<'_>,
    style: u64,
    variation: Variation,
) -> Result<()> {
    let (message, _) = style_payload(workspace, style)?;
    workspace.replace_payload(style, message, variation.payload)
}

fn derive_variation(
    workspace: &mut Workspace<'_>,
    context: &Context,
    variation: Variation,
) -> Result<()> {
    let identifier = workspace.next_object_identifier()?;
    let parent = variation.parent;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: MESSAGE_TYPE,
            data: variation.payload,
        }],
        workspace.limits(),
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = MESSAGE_VERSION.to_vec().into();
    info.object_references.push(parent);
    repoint_drawable(workspace, context.drawable, context.style, identifier)?;
    register(workspace, context.stylesheet, context.style, identifier)?;
    workspace.insert(context.style_location, object)?;
    let style_member = workspace.member(context.style_location).to_owned();
    let drawable_member = workspace.member(context.drawable_location).to_owned();
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        workspace.add_object_uuids(style_component, &[identifier])?;
        if let Some(drawable_component) = workspace.component_for_member(&drawable_member)?
            && drawable_component != style_component
        {
            workspace.add_external_reference(drawable_component, style_component, identifier)?;
        }
    }
    workspace.set_last_object_identifier(identifier)
}

fn collapse(workspace: &mut Workspace<'_>, context: &Context, parent: u64) -> Result<()> {
    repoint_drawable(workspace, context.drawable, context.style, parent)?;
    unregister(workspace, context.stylesheet, parent, context.style)?;
    workspace.remove(context.style)?;
    let style_member = workspace.member(context.style_location).to_owned();
    let drawable_member = workspace.member(context.drawable_location).to_owned();
    if let Some(style_component) = workspace.component_for_member(&style_member)? {
        workspace.remove_object_uuids(style_component, &[context.style])?;
        workspace.remove_external_references_to(style_component, context.style)?;
        if let Some(drawable_component) = workspace.component_for_member(&drawable_member)?
            && drawable_component != style_component
        {
            workspace.add_external_reference(drawable_component, style_component, parent)?;
        }
    }
    workspace.release_identifier_suffix(&[context.style])
}

impl Styles for Workspace<'_> {
    fn shape_style(&self, identifier: u64) -> Option<&[u8]> {
        style_payload(self, identifier)
            .ok()
            .map(|(_, payload)| payload)
    }
}

/// The planning target of shape drawable `drawable` in a workspace: its
/// style and whether the shape alone uses it.
///
/// # Errors
///
/// An invalid shape or style graph.
pub fn target(workspace: &Workspace<'_>, drawable: u64) -> Result<Target> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    Ok(Target::new(style, is_exclusive(workspace, style)?))
}

/// Publish `plan` (made for shape drawable `drawable`'s current style) in a
/// workspace: the steps the module documentation lists.
///
/// # Errors
///
/// An invalid shape, style, stylesheet or registry.
pub fn apply(workspace: &mut Workspace<'_>, drawable: u64, plan: Plan) -> Result<()> {
    let (_, style) = shape_style_reference(workspace, drawable)?;
    match plan {
        Plan::Unchanged => Ok(()),
        Plan::Rewrite(variation) => {
            context(workspace, drawable, style)?;
            rewrite_variation(workspace, style, variation)
        },
        Plan::Derive(variation) => {
            let context = context(workspace, drawable, style)?;
            derive_variation(workspace, &context, variation)
        },
        Plan::Collapse { parent, .. } => {
            let context = context(workspace, drawable, style)?;
            collapse(workspace, &context, parent)
        },
    }
}

fn slots_of<P: ShapeProperty>(property: &P) -> Vec<Slot> {
    property
        .groups()
        .iter()
        .flat_map(|group| group.iter().copied())
        .collect()
}

/// Give shape drawable `drawable` its own `value` of `property`. Returns
/// `false` when the shape already shows it.
///
/// # Errors
///
/// An invalid shape, style, stylesheet or registry; a value the property
/// cannot store; a write that does not read back.
pub fn set<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
    value: &P::Value,
) -> Result<bool> {
    let target = target(workspace, drawable)?;
    if inherited(&*workspace, target.style(), property)? == *value {
        return Ok(false);
    }
    let writes = property.writes(value)?;
    let (_, payload) = style_payload(workspace, target.style())?;
    if target.exclusive() {
        for (slot, _) in &writes {
            if let Some(stored) = stored_slot(payload, *slot)? {
                property.replaceable(*slot, stored)?;
            }
        }
    }
    let plan = plan_set(&*workspace, target, &writes)?;
    apply(workspace, drawable, plan)?;
    if self::value(workspace, drawable, property)? != *value {
        return Err(Error::Invalid("a shape-style write does not read back"));
    }
    Ok(true)
}

/// Remove shape drawable `drawable`'s own value of `property`, restoring
/// the inherited one. Returns `false` when its style defines none (or, for
/// a lenient property, is not a minimal variation).
///
/// # Errors
///
/// The [`set`] errors; a style with the property and unsupported direct
/// overrides.
pub fn reset<P: ShapeProperty>(
    workspace: &mut Workspace<'_>,
    drawable: u64,
    property: &P,
) -> Result<bool> {
    let target = target(workspace, drawable)?;
    let slots = slots_of(property);
    let (_, payload) = style_payload(workspace, target.style())?;
    if property.lenient_reset() && direct_overrides(payload)?.is_none() {
        return Ok(false);
    }
    let mut defined = false;
    for slot in &slots {
        if let Some(stored) = stored_slot(payload, *slot)? {
            property.replaceable(*slot, stored)?;
            defined = true;
        }
    }
    if !defined {
        return Ok(false);
    }
    let parent = style_parts(payload)?
        .parent
        .ok_or(Error::Invalid("a shape-style variation has no parent"))?;
    let inherited_value = inherited(&*workspace, parent, property)?;
    let fallback = property
        .writes(&inherited_value)?
        .into_iter()
        .filter_map(|(slot, value)| value.map(|value| (slot, value)))
        .collect::<Vec<_>>();
    let plan = plan_reset(&*workspace, target, &slots, &fallback)?;
    apply(workspace, drawable, plan)?;
    if value(workspace, drawable, property)? != inherited_value {
        return Err(Error::Invalid("a shape-style reset does not read back"));
    }
    Ok(true)
}

/// The fill, the stroke, the shadow, the effects (opacity and reflection),
/// the frame text layout and the text columns.
pub mod property {
    use litchi_iwa_common::color::Rgba;
    use litchi_iwa_common::shape::effects::{self, Effects, Reflection, ReflectionOpacity};
    use litchi_iwa_common::shape::fill::Fill;
    use litchi_iwa_common::shape::shadow::{
        Angle, Appearance, BlurRadius, Contact, Curve, Curved, Drop, Offset, Opacity, Perspective,
        Shadow,
    };
    use litchi_iwa_common::shape::stroke::Stroke;
    use litchi_iwa_common::text::layout::{AutoSize, Inset, Insets, Layout, VerticalAlignment};
    use litchi_iwa_protos::chart_style_codec::Limits as CodecLimits;
    use litchi_iwa_protos::shape_style_codec::native as stored;
    use litchi_iwa_text::columns::{Columns, Count, Following, Gap, Variable, Width};

    use super::{ShapeProperty, Slot, Value};
    use crate::drawing;
    use crate::graph::{self, Error, Result, push_message_field, push_varint_field};

    fn first<'a>(stored: &[Option<&'a [u8]>]) -> Option<&'a [u8]> {
        stored.first().copied().flatten()
    }

    /// The shape fill (`TSD.ShapeStylePropertiesArchive.fill`). Image fills
    /// move data references and media and are refused here.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Fills;

    impl ShapeProperty for Fills {
        type Value = Fill;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[&[Slot::Fill]]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Fill> {
            let Some(payload) = first(stored) else {
                return Ok(Fill::None);
            };
            drawing::fill(&stored::fill(payload, CodecLimits::DEFAULT)?)
                .map_err(|_| Error::Invalid("an unsupported stored fill"))
        }

        fn writes(&self, value: &Fill) -> Result<Vec<(Slot, Option<Value>)>> {
            if matches!(value, Fill::Image(_)) {
                return Err(Error::Invalid(
                    "image fills are written by the format families (data references and media)",
                ));
            }
            Ok(vec![(
                Slot::Fill,
                Some(Value::message(drawing::encode_fill(value))),
            )])
        }

        fn replaceable(&self, _slot: Slot, payload: &[u8]) -> Result<()> {
            if matches!(self.read(&[Some(payload)])?, Fill::Image(_)) {
                return Err(Error::Invalid(
                    "image fills are released by the format families (data references and media)",
                ));
            }
            Ok(())
        }
    }

    /// The shape stroke (`TSD.ShapeStylePropertiesArchive.stroke`; `None`
    /// is the hidden stroke).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Strokes;

    impl ShapeProperty for Strokes {
        type Value = Option<Stroke>;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[&[Slot::Stroke]]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Option<Stroke>> {
            let Some(payload) = first(stored) else {
                return Ok(None);
            };
            drawing::stroke(&stored::stroke(payload, CodecLimits::DEFAULT)?)
                .map_err(|_| Error::Invalid("an unsupported stored stroke"))
        }

        fn writes(&self, value: &Option<Stroke>) -> Result<Vec<(Slot, Option<Value>)>> {
            Ok(vec![(
                Slot::Stroke,
                Some(Value::message(drawing::encode_stroke(*value))),
            )])
        }
    }

    fn varint_of(bytes: &[u8]) -> Result<u64> {
        litchi_iwa_common::decode_varint_from_bytes(bytes)
            .map(|(value, _)| value)
            .map_err(|_| Error::Invalid("a stored varint is malformed"))
    }

    fn null_flag(bytes: Option<&[u8]>) -> Result<Option<bool>> {
        bytes
            .map(|bytes| match varint_of(bytes)? {
                0 => Ok(false),
                1 => Ok(true),
                _ => Err(Error::Invalid("a stored bool is not 0 or 1")),
            })
            .transpose()
    }

    fn float_of(bytes: &[u8], number: u32) -> Result<Option<f32>> {
        let mut found = None;
        for field in graph::fields(bytes)? {
            if field.number() != number {
                continue;
            }
            if field.wire_type() != 5 || found.is_some() {
                return Err(Error::Invalid("a stored float is invalid or repeated"));
            }
            let raw: [u8; 4] = bytes[field.key_end()..field.end()]
                .try_into()
                .map_err(|_| Error::Invalid("a stored float is truncated"))?;
            found = Some(f32::from_le_bytes(raw));
        }
        Ok(found)
    }

    fn push_float(output: &mut Vec<u8>, number: u32, value: f32) {
        graph::push_varint(output, (u64::from(number) << 3) | 5);
        output.extend_from_slice(&value.to_le_bytes());
    }

    const VERTICAL_ALIGNMENTS: [VerticalAlignment; 4] = [
        VerticalAlignment::Top,
        VerticalAlignment::Middle,
        VerticalAlignment::Bottom,
        VerticalAlignment::Justified,
    ];

    fn vertical_alignment_native(value: VerticalAlignment) -> i32 {
        match value {
            VerticalAlignment::Middle => 1,
            VerticalAlignment::Bottom => 2,
            VerticalAlignment::Justified => 3,
            _ => 0,
        }
    }

    fn insets_of(padding: &[u8]) -> Result<Insets> {
        let side = |number| -> Result<Inset> {
            Inset::from_points(float_of(padding, number)?.unwrap_or(0.0))
                .map_err(|_| Error::Invalid("a stored text inset is invalid"))
        };
        Ok(Insets::new(side(1)?, side(2)?, side(3)?, side(4)?))
    }

    fn encode_insets(insets: Insets) -> Vec<u8> {
        let mut padding = Vec::with_capacity(20);
        push_float(&mut padding, 1, insets.left().points());
        push_float(&mut padding, 2, insets.top().points());
        push_float(&mut padding, 3, insets.right().points());
        push_float(&mut padding, 4, insets.bottom().points());
        padding
    }

    /// The frame text layout (`TSWP.ShapeStylePropertiesArchive`
    /// `shrink_to_fit`, `vertical_alignment`, `padding_null` and `padding`).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct TextLayout;

    impl ShapeProperty for TextLayout {
        type Value = Layout;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[
                &[Slot::ShrinkToFit],
                &[Slot::VerticalAlignment],
                &[Slot::PaddingNull, Slot::Padding],
            ]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Layout> {
            let [shrink, alignment, padding_null, padding] = stored else {
                return Err(Error::Invalid("a text layout needs four slots"));
            };
            let auto_size = match null_flag(*shrink)? {
                Some(true) => AutoSize::ShrinkToFit,
                _ => AutoSize::Fixed,
            };
            let vertical_alignment = match alignment {
                Some(bytes) => {
                    let native = varint_of(bytes)?;
                    *VERTICAL_ALIGNMENTS
                        .iter()
                        .find(|candidate| {
                            i64::from(vertical_alignment_native(**candidate)) as u64 == native
                        })
                        .ok_or(Error::Invalid("an unknown vertical text alignment"))?
                },
                None => VerticalAlignment::Top,
            };
            let insets = match (null_flag(*padding_null)?, padding) {
                (Some(true), None) => Insets::ZERO,
                (Some(true), Some(_)) => {
                    return Err(Error::Invalid("text padding is both null and populated"));
                },
                (Some(false), None) => {
                    return Err(Error::Invalid("text padding is non-null but missing"));
                },
                (_, Some(padding)) => insets_of(padding)?,
                (None, None) => Insets::ZERO,
            };
            Ok(Layout::new(vertical_alignment, insets, auto_size))
        }

        fn writes(&self, value: &Layout) -> Result<Vec<(Slot, Option<Value>)>> {
            Ok(vec![
                (
                    Slot::ShrinkToFit,
                    Some(Value::bool(matches!(
                        value.auto_size(),
                        AutoSize::ShrinkToFit
                    ))),
                ),
                (
                    Slot::VerticalAlignment,
                    Some(Value::int32(vertical_alignment_native(
                        value.vertical_alignment(),
                    ))),
                ),
                (Slot::PaddingNull, None),
                (
                    Slot::Padding,
                    Some(Value::message(encode_insets(value.insets()))),
                ),
            ])
        }

        fn lenient_reset(&self) -> bool {
            true
        }
    }

    fn columns_of(bytes: &[u8]) -> Result<Columns> {
        let invalid = |_| Error::Invalid("stored text columns are invalid");
        let equal = super::single(bytes, 1)?;
        let variable = super::single(bytes, 2)?;
        match (equal, variable) {
            (Some(equal), None) => {
                let count = graph::varint(equal, 1)?
                    .ok_or(Error::Invalid("equal text columns have no count"))?;
                let count = u32::try_from(count)
                    .map_err(|_| Error::Invalid("a column count exceeds u32"))?;
                Ok(Columns::equal(
                    Count::new(count).map_err(invalid)?,
                    float_of(equal, 2)?
                        .map(Gap::from_points)
                        .transpose()
                        .map_err(invalid)?,
                ))
            },
            (None, Some(variable)) => {
                let first = float_of(variable, 1)?
                    .ok_or(Error::Invalid("variable text columns have no first width"))?;
                let following = graph::repeated(variable, 2)?
                    .into_iter()
                    .map(|column| {
                        let gap =
                            float_of(column, 1)?.ok_or(Error::Invalid("a column has no gap"))?;
                        let width =
                            float_of(column, 2)?.ok_or(Error::Invalid("a column has no width"))?;
                        Ok(Following::new(
                            Gap::from_points(gap).map_err(invalid)?,
                            Width::from_points(width).map_err(invalid)?,
                        ))
                    })
                    .collect::<Result<Vec<_>>>()?;
                Variable::new(Width::from_points(first).map_err(invalid)?, following)
                    .map(Columns::Variable)
                    .map_err(invalid)
            },
            _ => Err(Error::Invalid("stored text columns have no single layout")),
        }
    }

    fn encode_columns(columns: &Columns) -> Vec<u8> {
        let mut output = Vec::new();
        match columns {
            Columns::Equal(equal) => {
                let mut inner = Vec::new();
                push_varint_field(&mut inner, 1, u64::from(equal.count().get()));
                if let Some(gap) = equal.gap() {
                    push_float(&mut inner, 2, gap.points());
                }
                push_message_field(&mut output, 1, &inner);
            },
            Columns::Variable(variable) => {
                let mut inner = Vec::new();
                push_float(&mut inner, 1, variable.first_width().points());
                for column in variable.following() {
                    let mut entry = Vec::with_capacity(10);
                    push_float(&mut entry, 1, column.gap().points());
                    push_float(&mut entry, 2, column.width().points());
                    push_message_field(&mut inner, 2, &entry);
                }
                push_message_field(&mut output, 2, &inner);
            },
        }
        output
    }

    /// The text columns of a shape's text frame (`columns_null` and
    /// `columns`).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct TextColumns;

    impl ShapeProperty for TextColumns {
        type Value = Columns;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[&[Slot::ColumnsNull, Slot::Columns]]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Columns> {
            let [columns_null, columns] = stored else {
                return Err(Error::Invalid("text columns need two slots"));
            };
            match (null_flag(*columns_null)?, columns) {
                (Some(true), None) | (None, None) => Ok(Columns::default()),
                (Some(true), Some(_)) => {
                    Err(Error::Invalid("text columns are both null and populated"))
                },
                (Some(false), None) => Err(Error::Invalid("text columns are non-null but missing")),
                (_, Some(columns)) => columns_of(columns),
            }
        }

        fn writes(&self, value: &Columns) -> Result<Vec<(Slot, Option<Value>)>> {
            Ok(vec![
                (Slot::ColumnsNull, None),
                (Slot::Columns, Some(Value::message(encode_columns(value)))),
            ])
        }

        fn lenient_reset(&self) -> bool {
            true
        }
    }

    fn varint_field(bytes: &[u8], number: u32) -> Result<Option<u64>> {
        graph::varint(bytes, number)
    }

    fn color_of(bytes: &[u8]) -> Result<Rgba> {
        drawing::color(&stored::color(bytes, CodecLimits::DEFAULT)?)
            .map_err(|_| Error::Invalid("an unsupported stored color"))
    }

    const FULL_TURN: f32 = 360.0;

    fn angle_of(native: f32) -> Result<Angle> {
        if !native.is_finite() || !(0.0..FULL_TURN).contains(&native) {
            return Err(Error::Invalid("a shadow angle is not canonical"));
        }
        Angle::from_degrees(if native == 0.0 {
            0.0
        } else {
            FULL_TURN - native
        })
        .map_err(|_| Error::Invalid("a shadow angle is invalid"))
    }

    fn angle_native(angle: Angle) -> f32 {
        if angle == Angle::ZERO {
            0.0
        } else {
            FULL_TURN - angle.degrees()
        }
    }

    fn shadow_of(bytes: &[u8]) -> Result<Shadow> {
        let invalid = |_| Error::Invalid("a stored shadow is invalid");
        let numbers = super::field_numbers(bytes)?;
        if numbers.is_empty() {
            return Ok(Shadow::Disabled);
        }
        if varint_field(bytes, 6)? == Some(0) {
            if numbers == [6] {
                return Ok(Shadow::Disabled);
            }
            return Err(Error::Invalid("a disabled shadow keeps enabled properties"));
        }
        let radius = match varint_field(bytes, 4)? {
            Some(raw) => u32::try_from(raw as i64)
                .map_err(|_| Error::Invalid("a shadow blur radius is negative"))?,
            None => 1,
        };
        let appearance = Appearance::new(
            match super::single(bytes, 1)? {
                Some(color) => color_of(color)?,
                None => Rgba::black(),
            },
            BlurRadius::from_points(radius).map_err(invalid)?,
            Offset::from_points(float_of(bytes, 3)?.unwrap_or(5.0)).map_err(invalid)?,
            Opacity::new(float_of(bytes, 5)?.unwrap_or(1.0)).map_err(invalid)?,
        );
        let drop = super::single(bytes, 8)?;
        let contact = super::single(bytes, 9)?;
        let curved = super::single(bytes, 10)?;
        let angle = float_of(bytes, 2)?.unwrap_or(315.0);
        match varint_field(bytes, 7)?.unwrap_or(0) {
            0 => {
                if contact.is_some() || curved.is_some() {
                    return Err(Error::Invalid("a shadow combines exclusive families"));
                }
                Ok(Shadow::Drop(Drop::new(appearance, angle_of(angle)?)))
            },
            1 => {
                if drop.is_some() || curved.is_some() {
                    return Err(Error::Invalid("a shadow combines exclusive families"));
                }
                let height = match contact {
                    Some(contact) => float_of(contact, 2)?.unwrap_or(0.2),
                    None => 0.2,
                };
                let offset = match contact {
                    Some(contact) => float_of(contact, 4)?.unwrap_or(0.0),
                    None => 0.0,
                };
                Ok(Shadow::Contact(
                    Contact::new(
                        appearance,
                        Perspective::from_height(height).map_err(invalid)?,
                    )
                    .with_contact_offset(Offset::from_points(offset).map_err(invalid)?),
                ))
            },
            2 => {
                if drop.is_some() || contact.is_some() {
                    return Err(Error::Invalid("a shadow combines exclusive families"));
                }
                let curve = match curved {
                    Some(curved) => float_of(curved, 1)?.unwrap_or(0.6),
                    None => 0.6,
                };
                Ok(Shadow::Curved(Curved::new(
                    appearance,
                    angle_of(angle)?,
                    Curve::new(curve).map_err(invalid)?,
                )))
            },
            _ => Err(Error::Invalid("an unknown shadow type")),
        }
    }

    fn encode_shadow(shadow: Shadow) -> Vec<u8> {
        let mut output = Vec::new();
        let (appearance, angle, kind) = match shadow {
            Shadow::Drop(drop) => (drop.appearance(), angle_native(drop.angle()), 0),
            Shadow::Contact(contact) => (contact.appearance(), 0.0, 1),
            Shadow::Curved(curved) => (curved.appearance(), angle_native(curved.angle()), 2),
            _ => return output,
        };
        push_message_field(&mut output, 1, &drawing::encode_color(appearance.color()));
        push_float(&mut output, 2, angle);
        push_float(&mut output, 3, appearance.offset().points());
        push_varint_field(&mut output, 4, u64::from(appearance.blur_radius().points()));
        push_float(&mut output, 5, appearance.opacity().get());
        push_varint_field(&mut output, 6, 1);
        push_varint_field(&mut output, 7, kind);
        match shadow {
            Shadow::Contact(contact) => {
                let mut inner = Vec::new();
                push_float(&mut inner, 2, contact.perspective().height());
                if contact.contact_offset() != Offset::ZERO {
                    push_float(&mut inner, 4, contact.contact_offset().points());
                }
                push_message_field(&mut output, 9, &inner);
            },
            Shadow::Curved(curved) => {
                let mut inner = Vec::new();
                push_float(&mut inner, 1, curved.curve().get());
                push_message_field(&mut output, 10, &inner);
            },
            _ => {},
        }
        output
    }

    /// The shape shadow (`TSD.ShapeStylePropertiesArchive.shadow`).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct Shadows;

    impl ShapeProperty for Shadows {
        type Value = Shadow;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[&[Slot::Shadow]]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Shadow> {
            first(stored).map_or(Ok(Shadow::Disabled), shadow_of)
        }

        fn writes(&self, value: &Shadow) -> Result<Vec<(Slot, Option<Value>)>> {
            Ok(vec![(
                Slot::Shadow,
                Some(Value::message(encode_shadow(*value))),
            )])
        }
    }

    /// The shape effects: opacity and reflection.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    pub struct ShapeEffects;

    impl ShapeProperty for ShapeEffects {
        type Value = Effects;

        fn groups(&self) -> &'static [&'static [Slot]] {
            &[&[Slot::Opacity], &[Slot::Reflection]]
        }

        fn read(&self, stored: &[Option<&[u8]>]) -> Result<Effects> {
            let [opacity, reflection] = stored else {
                return Err(Error::Invalid("effects need two slots"));
            };
            let opacity = match opacity {
                Some(bytes) => {
                    let raw: [u8; 4] = (*bytes)
                        .try_into()
                        .map_err(|_| Error::Invalid("a stored opacity is truncated"))?;
                    effects::Opacity::new(f32::from_le_bytes(raw))
                        .map_err(|_| Error::Invalid("a stored opacity is invalid"))?
                },
                None => effects::Opacity::OPAQUE,
            };
            let reflection = match reflection {
                Some(bytes) => match float_of(bytes, 1)? {
                    Some(value) => Reflection::Enabled(
                        ReflectionOpacity::new(value)
                            .map_err(|_| Error::Invalid("a stored reflection is invalid"))?,
                    ),
                    None => Reflection::Disabled,
                },
                None => Reflection::Disabled,
            };
            Ok(Effects::new(opacity, reflection))
        }

        fn writes(&self, value: &Effects) -> Result<Vec<(Slot, Option<Value>)>> {
            let mut reflection = Vec::new();
            if let Reflection::Enabled(opacity) = value.reflection() {
                push_float(&mut reflection, 1, opacity.value());
            }
            Ok(vec![
                (Slot::Opacity, Some(Value::float(value.opacity().value()))),
                (Slot::Reflection, Some(Value::message(reflection))),
            ])
        }
    }
}
