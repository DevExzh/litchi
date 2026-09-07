//! Source-preserving lazy support for `KN.SlideNodeArchive` build caches.
//!
//! The native slide-node message has a large, repeated build graph.  This
//! seam intentionally projects only the six scalar cache fields touched by
//! the native media duplicate/remove transaction.  The handwritten parser
//! remains authoritative for framing, duplicate detection, canonical values,
//! and unknown spans; the private Buffa view is a bounded parity check.

use core::convert::TryFrom;

use super::{
    Budget, DecodeError, DecodeLimit, DecodeOptions, DecodeReport, Field, Parser, RewriteReport,
    canonical_bool, checked_add_output, emit_varint_field, ensure_output, reserve_output,
    rewrite_report, validate_input, varint_field_size,
};

use super::projection;

const BUILD_EVENT_COUNT_FIELD: u32 = 15;
const BUILD_EVENT_COUNT_CACHE_VERSION_FIELD: u32 = 26;
const BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD: u32 = 22;
const HAS_EXPLICIT_BUILDS_FIELD: u32 = 20;
const HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD: u32 = 27;
const HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD: u32 = 23;

const BUILD_EVENT_COUNT_NAME: &str = "KN.SlideNodeArchive.buildEventCount";
const BUILD_EVENT_COUNT_CACHE_VERSION_NAME: &str =
    "KN.SlideNodeArchive.buildEventCountCacheVersion";
const BUILD_EVENT_COUNT_IS_UP_TO_DATE_NAME: &str = "KN.SlideNodeArchive.buildEventCountIsUpToDate";
const HAS_EXPLICIT_BUILDS_NAME: &str = "KN.SlideNodeArchive.hasExplicitBuilds";
const HAS_EXPLICIT_BUILDS_CACHE_VERSION_NAME: &str =
    "KN.SlideNodeArchive.hasExplicitBuildsCacheVersion";
const HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_NAME: &str =
    "KN.SlideNodeArchive.hasExplicitBuildsIsUpToDate";

/// Borrowed scalar cache state from one complete native slide-node payload.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNodeBuildCacheSnapshot<'source> {
    source: &'source [u8],
    build_event_count: Option<u32>,
    build_event_count_cache_version: Option<u32>,
    build_event_count_is_up_to_date: Option<bool>,
    has_explicit_builds: Option<bool>,
    has_explicit_builds_cache_version: Option<u32>,
    has_explicit_builds_is_up_to_date: Option<bool>,
}

impl<'source> SlideNodeBuildCacheSnapshot<'source> {
    /// Return the exact validated source payload.
    #[must_use]
    pub const fn source(self) -> &'source [u8] {
        self.source
    }

    /// Return the optional native build-event count.
    #[must_use]
    pub const fn build_event_count(self) -> Option<u32> {
        self.build_event_count
    }

    /// Return the optional native build-event-count cache version.
    #[must_use]
    pub const fn build_event_count_cache_version(self) -> Option<u32> {
        self.build_event_count_cache_version
    }

    /// Return the deprecated native build-event-count validity flag.
    #[must_use]
    pub const fn build_event_count_is_up_to_date(self) -> Option<bool> {
        self.build_event_count_is_up_to_date
    }

    /// Return the optional native explicit-builds flag.
    #[must_use]
    pub const fn has_explicit_builds(self) -> Option<bool> {
        self.has_explicit_builds
    }

    /// Return the optional native explicit-builds cache version.
    #[must_use]
    pub const fn has_explicit_builds_cache_version(self) -> Option<u32> {
        self.has_explicit_builds_cache_version
    }

    /// Return the deprecated native explicit-builds validity flag.
    #[must_use]
    pub const fn has_explicit_builds_is_up_to_date(self) -> Option<bool> {
        self.has_explicit_builds_is_up_to_date
    }
}

/// Explicit request to invalidate the native slide-node build caches.
#[doc(hidden)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideNodeBuildCacheEdit;

impl SlideNodeBuildCacheEdit {
    /// Construct an explicit native cache invalidation.
    #[must_use]
    pub const fn invalidate() -> Self {
        Self
    }
}

/// A source-witnessed, bounded build-cache rewrite awaiting commit.
#[doc(hidden)]
#[derive(Debug)]
pub struct PreparedSlideNodeBuildCacheRewrite<'source> {
    source: &'source [u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
    snapshot: SlideNodeBuildCacheSnapshot<'source>,
    output_bytes: usize,
    estimated_work_bytes: usize,
}

impl<'source> PreparedSlideNodeBuildCacheRewrite<'source> {
    /// Return the exact source witness retained by this preparation.
    #[must_use]
    pub const fn source(&self) -> &'source [u8] {
        self.source
    }

    /// Return the output size calculated before allocation.
    #[must_use]
    pub const fn output_bytes(&self) -> usize {
        self.output_bytes
    }

    /// Return the validated scalar source snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> SlideNodeBuildCacheSnapshot<'source> {
        self.snapshot
    }

    /// Emit, read back, and return the candidate and bounded operation report.
    pub fn commit(self) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
        let mut output = reserve_output(self.output_bytes)?;
        let mut budget = Budget::new(self.source, self.options);
        emit_cache(
            self.source,
            self.edit,
            self.options,
            &mut budget,
            &mut output,
        )?;
        if output.len() != self.output_bytes {
            return Err(DecodeError::projection());
        }

        let readback_options = self
            .options
            .with_max_message_bytes(self.options.max_message_bytes().max(output.len()));
        let (readback, readback_report) =
            decode_slide_node_build_cache_with_report(&output, readback_options)?;
        if readback != expected_snapshot(&output, self.edit) {
            return Err(DecodeError::projection());
        }

        let changed = output.as_slice() != self.source;
        let mut report = rewrite_report(
            self.source,
            &budget,
            &readback_report,
            self.output_bytes,
            changed,
        );
        report.work_bytes = report.work_bytes.saturating_add(self.estimated_work_bytes);
        if report.work_bytes > self.options.max_work_bytes() {
            return Err(DecodeError::limited(DecodeLimit::Work {
                observed: report.work_bytes,
                maximum: self.options.max_work_bytes(),
            }));
        }
        Ok((output, report))
    }
}

/// Decode one native slide-node build-cache payload lazily.
#[doc(hidden)]
pub fn decode_slide_node_build_cache<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<SlideNodeBuildCacheSnapshot<'source>, DecodeError> {
    decode_slide_node_build_cache_with_report(source, options).map(|(snapshot, _)| snapshot)
}

/// Decode one native slide-node build-cache payload with exact resource usage.
#[doc(hidden)]
pub fn decode_slide_node_build_cache_with_report<'source>(
    source: &'source [u8],
    options: DecodeOptions,
) -> Result<(SlideNodeBuildCacheSnapshot<'source>, DecodeReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_cache(source, options, &mut budget)?;
    cross_check_projection(source, options, snapshot, &mut budget)?;
    Ok((snapshot, budget.report(source.len())))
}

/// Prepare a source-witnessed build-cache rewrite.
#[doc(hidden)]
pub fn prepare_slide_node_build_cache_rewrite<'source>(
    source: &'source [u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
) -> Result<PreparedSlideNodeBuildCacheRewrite<'source>, DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_cache(source, options, &mut budget)?;
    cross_check_projection(source, options, snapshot, &mut budget)?;
    let output_bytes = measure_cache(source, edit, options, &mut budget)?;
    ensure_output(output_bytes, options)?;
    let estimated_work_bytes = output_bytes
        .checked_mul(4)
        .and_then(|value| value.checked_add(budget.work_bytes))
        .ok_or_else(DecodeError::invalid)?;
    if estimated_work_bytes > options.max_work_bytes() {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: estimated_work_bytes,
            maximum: options.max_work_bytes(),
        }));
    }
    Ok(PreparedSlideNodeBuildCacheRewrite {
        source,
        edit,
        options,
        snapshot,
        output_bytes,
        estimated_work_bytes,
    })
}

/// Rewrite one native slide-node build-cache payload.
#[doc(hidden)]
pub fn rewrite_slide_node_build_cache(
    source: &[u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_slide_node_build_cache_with_report(source, edit, options).map(|(output, _)| output)
}

/// Rewrite one native slide-node build-cache payload with exact usage.
#[doc(hidden)]
pub fn rewrite_slide_node_build_cache_with_report(
    source: &[u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    prepare_slide_node_build_cache_rewrite(source, edit, options)?.commit()
}

/// Rewrite the four build-cache scalars after authoring a fresh build.
///
/// This follows the native Keynote writer's event-count behavior: a zero
/// count removes `buildEventCount`, stores `u32::MAX` in its cache version,
/// and writes `hasExplicitBuilds = false`; a nonzero count stores the count,
/// version `2`, and `hasExplicitBuilds = true`.  The source remains the
/// preservation authority, so unrelated fields and unknown spans are copied
/// byte-for-byte.  The existing cache invalidation edit intentionally keeps
/// its separate remove-operation semantics.
#[doc(hidden)]
pub fn rewrite_slide_node_build_cache_for_event_count(
    source: &[u8],
    event_count: u32,
    options: DecodeOptions,
) -> Result<Vec<u8>, DecodeError> {
    rewrite_slide_node_build_cache_for_event_count_with_report(source, event_count, options)
        .map(|(output, _)| output)
}

/// Rewrite build-cache scalars and return exact bounded resource evidence.
#[doc(hidden)]
pub fn rewrite_slide_node_build_cache_for_event_count_with_report(
    source: &[u8],
    event_count: u32,
    options: DecodeOptions,
) -> Result<(Vec<u8>, RewriteReport), DecodeError> {
    validate_input(source, options)?;
    let mut budget = Budget::new(source, options);
    let snapshot = parse_cache(source, options, &mut budget)?;
    cross_check_projection(source, options, snapshot, &mut budget)?;
    let values = expected_event_count_values(event_count, snapshot);
    let output_bytes = measure_cache_values(source, values, options, &mut budget)?;
    ensure_output(output_bytes, options)?;
    let estimated_work_bytes = output_bytes
        .checked_mul(4)
        .and_then(|value| value.checked_add(budget.work_bytes))
        .ok_or_else(DecodeError::invalid)?;
    if estimated_work_bytes > options.max_work_bytes() {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: estimated_work_bytes,
            maximum: options.max_work_bytes(),
        }));
    }
    let mut output = reserve_output(output_bytes)?;
    emit_cache_values(source, values, options, &mut budget, &mut output)?;
    if output.len() != output_bytes {
        return Err(DecodeError::projection());
    }

    let readback_options =
        options.with_max_message_bytes(options.max_message_bytes().max(output.len()));
    let (readback, readback_report) =
        decode_slide_node_build_cache_with_report(&output, readback_options)?;
    if CacheValues::from_snapshot(readback) != values {
        return Err(DecodeError::projection());
    }
    let changed = output.as_slice() != source;
    let mut report = rewrite_report(source, &budget, &readback_report, output_bytes, changed);
    report.work_bytes = report.work_bytes.saturating_add(estimated_work_bytes);
    if report.work_bytes > options.max_work_bytes() {
        return Err(DecodeError::limited(DecodeLimit::Work {
            observed: report.work_bytes,
            maximum: options.max_work_bytes(),
        }));
    }
    Ok((output, report))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CacheValues {
    build_event_count: Option<u32>,
    build_event_count_cache_version: Option<u32>,
    build_event_count_is_up_to_date: Option<bool>,
    has_explicit_builds: Option<bool>,
    has_explicit_builds_cache_version: Option<u32>,
    has_explicit_builds_is_up_to_date: Option<bool>,
}

impl CacheValues {
    const fn from_snapshot(snapshot: SlideNodeBuildCacheSnapshot<'_>) -> Self {
        Self {
            build_event_count: snapshot.build_event_count,
            build_event_count_cache_version: snapshot.build_event_count_cache_version,
            build_event_count_is_up_to_date: snapshot.build_event_count_is_up_to_date,
            has_explicit_builds: snapshot.has_explicit_builds,
            has_explicit_builds_cache_version: snapshot.has_explicit_builds_cache_version,
            has_explicit_builds_is_up_to_date: snapshot.has_explicit_builds_is_up_to_date,
        }
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct Presence {
    build_event_count: bool,
    build_event_count_cache_version: bool,
    build_event_count_is_up_to_date: bool,
    has_explicit_builds: bool,
    has_explicit_builds_cache_version: bool,
    has_explicit_builds_is_up_to_date: bool,
}

fn expected_values(_edit: SlideNodeBuildCacheEdit) -> CacheValues {
    CacheValues {
        build_event_count: None,
        build_event_count_cache_version: Some(u32::MAX),
        build_event_count_is_up_to_date: None,
        has_explicit_builds: None,
        has_explicit_builds_cache_version: Some(u32::MAX),
        has_explicit_builds_is_up_to_date: None,
    }
}

fn expected_event_count_values(
    event_count: u32,
    snapshot: SlideNodeBuildCacheSnapshot<'_>,
) -> CacheValues {
    CacheValues {
        build_event_count: (event_count != 0).then_some(event_count),
        build_event_count_cache_version: Some(if event_count == 0 { u32::MAX } else { 2 }),
        build_event_count_is_up_to_date: snapshot.build_event_count_is_up_to_date,
        has_explicit_builds: Some(event_count != 0),
        has_explicit_builds_cache_version: Some(2),
        has_explicit_builds_is_up_to_date: snapshot.has_explicit_builds_is_up_to_date,
    }
}

fn expected_snapshot<'source>(
    source: &'source [u8],
    edit: SlideNodeBuildCacheEdit,
) -> SlideNodeBuildCacheSnapshot<'source> {
    let values = expected_values(edit);
    SlideNodeBuildCacheSnapshot {
        source,
        build_event_count: values.build_event_count,
        build_event_count_cache_version: values.build_event_count_cache_version,
        build_event_count_is_up_to_date: values.build_event_count_is_up_to_date,
        has_explicit_builds: values.has_explicit_builds,
        has_explicit_builds_cache_version: values.has_explicit_builds_cache_version,
        has_explicit_builds_is_up_to_date: values.has_explicit_builds_is_up_to_date,
    }
}

fn parse_cache<'source>(
    source: &'source [u8],
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<SlideNodeBuildCacheSnapshot<'source>, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut values = CacheValues {
        build_event_count: None,
        build_event_count_cache_version: None,
        build_event_count_is_up_to_date: None,
        has_explicit_builds: None,
        has_explicit_builds_cache_version: None,
        has_explicit_builds_is_up_to_date: None,
    };

    while let Some(field) = parser.next()? {
        match field.number {
            BUILD_EVENT_COUNT_FIELD => {
                if values.build_event_count.is_some() {
                    return Err(DecodeError::duplicate(BUILD_EVENT_COUNT_NAME));
                }
                values.build_event_count = Some(uint32(
                    field.varint(BUILD_EVENT_COUNT_NAME)?,
                    BUILD_EVENT_COUNT_NAME,
                )?);
            },
            BUILD_EVENT_COUNT_CACHE_VERSION_FIELD => {
                if values.build_event_count_cache_version.is_some() {
                    return Err(DecodeError::duplicate(BUILD_EVENT_COUNT_CACHE_VERSION_NAME));
                }
                values.build_event_count_cache_version = Some(uint32(
                    field.varint(BUILD_EVENT_COUNT_CACHE_VERSION_NAME)?,
                    BUILD_EVENT_COUNT_CACHE_VERSION_NAME,
                )?);
            },
            BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD => {
                if values.build_event_count_is_up_to_date.is_some() {
                    return Err(DecodeError::duplicate(BUILD_EVENT_COUNT_IS_UP_TO_DATE_NAME));
                }
                values.build_event_count_is_up_to_date = Some(canonical_bool(
                    field.varint(BUILD_EVENT_COUNT_IS_UP_TO_DATE_NAME)?,
                )?);
            },
            HAS_EXPLICIT_BUILDS_FIELD => {
                if values.has_explicit_builds.is_some() {
                    return Err(DecodeError::duplicate(HAS_EXPLICIT_BUILDS_NAME));
                }
                values.has_explicit_builds =
                    Some(canonical_bool(field.varint(HAS_EXPLICIT_BUILDS_NAME)?)?);
            },
            HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD => {
                if values.has_explicit_builds_cache_version.is_some() {
                    return Err(DecodeError::duplicate(
                        HAS_EXPLICIT_BUILDS_CACHE_VERSION_NAME,
                    ));
                }
                values.has_explicit_builds_cache_version = Some(uint32(
                    field.varint(HAS_EXPLICIT_BUILDS_CACHE_VERSION_NAME)?,
                    HAS_EXPLICIT_BUILDS_CACHE_VERSION_NAME,
                )?);
            },
            HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD => {
                if values.has_explicit_builds_is_up_to_date.is_some() {
                    return Err(DecodeError::duplicate(
                        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_NAME,
                    ));
                }
                values.has_explicit_builds_is_up_to_date = Some(canonical_bool(
                    field.varint(HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_NAME)?,
                )?);
            },
            _ => {},
        }
    }

    Ok(SlideNodeBuildCacheSnapshot {
        source,
        build_event_count: values.build_event_count,
        build_event_count_cache_version: values.build_event_count_cache_version,
        build_event_count_is_up_to_date: values.build_event_count_is_up_to_date,
        has_explicit_builds: values.has_explicit_builds,
        has_explicit_builds_cache_version: values.has_explicit_builds_cache_version,
        has_explicit_builds_is_up_to_date: values.has_explicit_builds_is_up_to_date,
    })
}

fn uint32(value: u64, _field: &'static str) -> Result<u32, DecodeError> {
    u32::try_from(value).map_err(|_| DecodeError::noncanonical("uint32 scalar exceeds uint32"))
}

fn cross_check_projection(
    source: &[u8],
    options: DecodeOptions,
    snapshot: SlideNodeBuildCacheSnapshot<'_>,
    budget: &mut Budget,
) -> Result<(), DecodeError> {
    let view: projection::SlideNodeBuildCacheArchiveLazyView<'_> = options
        .buffa()
        .decode_lazy_view(source)
        .map_err(DecodeError::from)?;
    budget.charge_work(source.len())?;
    let expected = CacheValues::from_snapshot(snapshot);
    let projected = CacheValues {
        build_event_count: view.build_event_count,
        build_event_count_cache_version: view.build_event_count_cache_version,
        build_event_count_is_up_to_date: view.build_event_count_is_up_to_date,
        has_explicit_builds: view.has_explicit_builds,
        has_explicit_builds_cache_version: view.has_explicit_builds_cache_version,
        has_explicit_builds_is_up_to_date: view.has_explicit_builds_is_up_to_date,
    };
    if projected != expected {
        return Err(DecodeError::projection());
    }
    Ok(())
}

fn replacement_for(number: u32, values: CacheValues) -> Option<Option<u64>> {
    match number {
        BUILD_EVENT_COUNT_FIELD => Some(values.build_event_count.map(u64::from)),
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD => {
            Some(values.build_event_count_cache_version.map(u64::from))
        },
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD => {
            Some(values.build_event_count_is_up_to_date.map(u64::from))
        },
        HAS_EXPLICIT_BUILDS_FIELD => Some(values.has_explicit_builds.map(u64::from)),
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD => {
            Some(values.has_explicit_builds_cache_version.map(u64::from))
        },
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD => {
            Some(values.has_explicit_builds_is_up_to_date.map(u64::from))
        },
        _ => None,
    }
}

fn field_name(number: u32) -> &'static str {
    match number {
        BUILD_EVENT_COUNT_FIELD => BUILD_EVENT_COUNT_NAME,
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD => BUILD_EVENT_COUNT_CACHE_VERSION_NAME,
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD => BUILD_EVENT_COUNT_IS_UP_TO_DATE_NAME,
        HAS_EXPLICIT_BUILDS_FIELD => HAS_EXPLICIT_BUILDS_NAME,
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD => HAS_EXPLICIT_BUILDS_CACHE_VERSION_NAME,
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD => HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_NAME,
        _ => "KN.SlideNodeArchive.unknown",
    }
}

fn mark_seen(presence: &mut Presence, number: u32) {
    match number {
        BUILD_EVENT_COUNT_FIELD => presence.build_event_count = true,
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD => {
            presence.build_event_count_cache_version = true;
        },
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD => {
            presence.build_event_count_is_up_to_date = true;
        },
        HAS_EXPLICIT_BUILDS_FIELD => presence.has_explicit_builds = true,
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD => {
            presence.has_explicit_builds_cache_version = true;
        },
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD => {
            presence.has_explicit_builds_is_up_to_date = true;
        },
        _ => {},
    }
}

fn was_seen(presence: Presence, number: u32) -> bool {
    match number {
        BUILD_EVENT_COUNT_FIELD => presence.build_event_count,
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD => presence.build_event_count_cache_version,
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD => presence.build_event_count_is_up_to_date,
        HAS_EXPLICIT_BUILDS_FIELD => presence.has_explicit_builds,
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD => presence.has_explicit_builds_cache_version,
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD => presence.has_explicit_builds_is_up_to_date,
        _ => false,
    }
}

fn cache_field_size(field: Field<'_>, values: CacheValues) -> Result<usize, DecodeError> {
    let replacement = replacement_for(field.number, values).ok_or_else(DecodeError::invalid)?;
    let current = field.varint(field_name(field.number))?;
    Ok(match replacement {
        Some(value) if value == current => field.raw.len(),
        Some(value) => varint_field_size(field.number, value),
        None => 0,
    })
}

fn measure_cache(
    source: &[u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    measure_cache_values(source, expected_values(edit), options, budget)
}

fn measure_cache_values(
    source: &[u8],
    values: CacheValues,
    options: DecodeOptions,
    budget: &mut Budget,
) -> Result<usize, DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut presence = Presence::default();
    let mut output = 0usize;
    while let Some(field) = parser.next()? {
        if replacement_for(field.number, values).is_some() {
            mark_seen(&mut presence, field.number);
            output = checked_add_output(output, cache_field_size(field, values)?, options)?;
        } else {
            output = checked_add_output(output, field.raw.len(), options)?;
        }
    }
    for number in [
        BUILD_EVENT_COUNT_FIELD,
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD,
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD,
        HAS_EXPLICIT_BUILDS_FIELD,
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD,
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD,
    ] {
        if was_seen(presence, number) {
            continue;
        }
        if let Some(value) = replacement_for(number, values).flatten() {
            output = checked_add_output(output, varint_field_size(number, value), options)?;
        }
    }
    Ok(output)
}

fn emit_cache(
    source: &[u8],
    edit: SlideNodeBuildCacheEdit,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    emit_cache_values(source, expected_values(edit), options, budget, output)
}

fn emit_cache_values(
    source: &[u8],
    values: CacheValues,
    options: DecodeOptions,
    budget: &mut Budget,
    output: &mut Vec<u8>,
) -> Result<(), DecodeError> {
    let mut parser = Parser::new(source, 1, options, budget)?;
    let mut presence = Presence::default();
    while let Some(field) = parser.next()? {
        let Some(replacement) = replacement_for(field.number, values) else {
            output.extend_from_slice(field.raw);
            continue;
        };
        mark_seen(&mut presence, field.number);
        let current = field.varint(field_name(field.number))?;
        match replacement {
            Some(value) if value == current => output.extend_from_slice(field.raw),
            Some(value) => emit_varint_field(field.number, value, output),
            None => {},
        }
    }
    for number in [
        BUILD_EVENT_COUNT_FIELD,
        BUILD_EVENT_COUNT_CACHE_VERSION_FIELD,
        BUILD_EVENT_COUNT_IS_UP_TO_DATE_FIELD,
        HAS_EXPLICIT_BUILDS_FIELD,
        HAS_EXPLICIT_BUILDS_CACHE_VERSION_FIELD,
        HAS_EXPLICIT_BUILDS_IS_UP_TO_DATE_FIELD,
    ] {
        if was_seen(presence, number) {
            continue;
        }
        if let Some(value) = replacement_for(number, values).flatten() {
            emit_varint_field(number, value, output);
        }
    }
    Ok(())
}

#[cfg(test)]
#[path = "node_cache_tests.rs"]
mod tests;
