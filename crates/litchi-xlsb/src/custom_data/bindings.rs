//! Bounded, source-preserving bindings for `BrtBeginExtConn14`.
//!
//! This module is deliberately independent of the XLSB package graph.  It
//! lends all source bytes from the caller, scans the External Data
//! Connections envelope once, and records only the spans needed by a Custom
//! Data transaction.  A rewrite is a single ordered splice over that source;
//! it does not re-emit unknown records or search the source for UTF-16 text.
//!
//! The admitted wire profile is the one pinned by
//! `docs/report/spec-gap-validation-evidence/xlsb-extconn14-grammar`: a
//! `BrtBeginExtConn14` record has a reserved `FRTBlank`, an `irstCulture`
//! `XLWideString`, and an `irstClientCubeUrn` `XLWideString`, in that order.
//! The record is effective only when it is a direct child of the known
//! External Data Connections `FRTEXTCONNECTIONS` wrapper.  Alternate-content
//! blocks, malformed or nested future-record wrappers, and direct records in
//! an unknown context are retained as opaque provenance and block identity
//! mutation when they could hide a reference.  The `FRTBlank` bytes are
//! retained as source bytes and ignored as required by the record grammar.

use std::fmt;
use std::mem::size_of;

/// `BrtBeginExtConn14` record kind (MS-XLSB 2.4.78).
const BEGIN_EXT_CONN14: u16 = 1068;
/// `BrtEndExtConn14` record kind.
const END_EXT_CONN14: u16 = 1069;
/// `BrtBeginExtConn15` record kind.  Its `irstId` is not a Custom Data UID.
const BEGIN_EXT_CONN15: u16 = 2109;
/// `BrtEndExtConn15` record kind.
const END_EXT_CONN15: u16 = 2110;
/// `BrtBeginExtConnections` record kind.
const BEGIN_EXT_CONNECTIONS: u16 = 429;
/// `BrtEndExtConnections` record kind.
const END_EXT_CONNECTIONS: u16 = 430;
/// `BrtBeginExtConnection` record kind.
const BEGIN_EXT_CONNECTION: u16 = 201;
/// `BrtEndExtConnection` record kind.
const END_EXT_CONNECTION: u16 = 202;
/// `BrtBeginECDbProps` record kind.
const BEGIN_EC_DB_PROPS: u16 = 203;
/// `BrtEndECDbProps` record kind.
const END_EC_DB_PROPS: u16 = 204;
/// `BrtFRTBegin` record kind.
const FRT_BEGIN: u16 = 35;
/// `BrtFRTEnd` record kind.
const FRT_END: u16 = 36;
/// `BrtACBegin` record kind.
const AC_BEGIN: u16 = 37;
/// `BrtACEnd` record kind.
const AC_END: u16 = 38;

/// The `DBType` value required by the ExtConn14 profile (`DBTOLEDB`).
const DBTOLEDB: u32 = 5;
/// The `CmdType` value required by a non-empty ExtConn14 collection
/// (`CMDCUBE`).
const CMDCUBE: u32 = 1;

const MAX_WIRE_PAYLOAD: usize = 0x0fff_ffff;
const HARD_CULTURE_UNITS: usize = 84; // "less than 85 characters"
const HARD_CLIENT_CUBE_URN_UNITS: usize = 65_535; // "less than 65536"

/// A half-open byte range in the borrowed connections source.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct Span {
    start: usize,
    end: usize,
}

impl Span {
    /// Construct a checked half-open range.
    const fn new(start: usize, end: usize) -> Result<Self> {
        if start <= end {
            Ok(Self { start, end })
        } else {
            Err(BindingError::InvalidSpan { start, end })
        }
    }

    /// Start offset, inclusive.
    #[must_use]
    const fn start(self) -> usize {
        self.start
    }

    /// End offset, exclusive.
    #[must_use]
    const fn end(self) -> usize {
        self.end
    }

    /// Number of bytes in the range.
    #[must_use]
    const fn len(self) -> usize {
        self.end - self.start
    }

    fn contains(self, other: Self) -> bool {
        self.start <= other.start && other.end <= self.end
    }

    fn overlaps(self, other: Self) -> bool {
        self.start < other.end && other.start < self.end
    }
}

/// A decoded BIFF12 record envelope and all of its source framing spans.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct RecordSpan {
    /// Complete record range, including kind and payload-length varints.
    range: Span,
    /// Payload range, excluding both header varints.
    payload: Span,
    /// Numeric BIFF12 record kind.
    kind: u16,
    /// Number of bytes in the kind varint.
    kind_len: u8,
    /// Number of bytes in the payload-length varint.
    length_len: u8,
}

/// Source spans for one `XLWideString`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct WideStringSpan {
    /// Complete field range, including the four-byte unit count.
    field: Span,
    /// Four-byte little-endian UTF-16 unit-count range.
    length: Span,
    /// UTF-16LE data range.
    data: Span,
    /// Decoded UTF-16 code-unit count.
    units: usize,
}

/// The validated `FRTProductVersion` on an admitted wrapper.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct FrtProductVersion {
    /// Application version from the first two bytes.
    version: u16,
    /// Fifteen-bit application identifier from the second two bytes.
    product: u16,
}

/// The complete source span of a known `FRTEXTCONNECTIONS` wrapper.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct FrtWrapperSpan {
    /// `BrtFRTBegin` record.
    begin: RecordSpan,
    /// `BrtFRTEnd` record.
    end: RecordSpan,
    /// Complete wrapper range.
    range: Span,
    /// The validated product/version header bytes.
    profile: FrtProductVersion,
}

/// The complete `BrtBeginExtConnections`/`BrtEndExtConnections` root span.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct ConnectionsRootSpan {
    /// Root begin record.
    begin: RecordSpan,
    /// Root end record.
    end: RecordSpan,
    /// Complete root range.
    range: Span,
}

/// Why a source range was kept opaque instead of becoming an effective edge.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum OpaqueReason {
    /// A possible reference occurred in `BrtACBegin`/`BrtACEnd`.
    AlternateContent,
    /// A possible reference occurred in an unsupported or nested FRT block.
    FutureRecordWrapper,
    /// ExtConn14 was outside the admitted wrapper/context profile.
    UnsupportedExtConn14,
    /// The required DBTOLEDB/CMDCUBE/PivotCache admission proof was absent.
    Context,
    /// A possible Custom Data reference was nested in ExtConn15.
    ExtConn15,
}

/// A balanced opaque source block that may prevent identity mutation.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct OpaqueBlock {
    /// Complete source range retained by the host.
    range: Span,
    /// Classification used for typed mutation refusal.
    reason: OpaqueReason,
    /// Whether a candidate Custom Data reference was observed in this block.
    contains_reference_candidate: bool,
    /// FRT product/version header when this is a future-record block.
    frt_profile: Option<FrtProductVersion>,
    /// AC product/version list when this is an alternate-content block.
    alternate_content_versions: Vec<AcProductVersion>,
}

/// One `ACProductVersion` from an alternate-content wrapper.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct AcProductVersion {
    /// Application version threshold.
    file_version: u16,
    /// Fifteen-bit application identifier.
    file_product: u16,
    /// Whether versions greater than or equal to `file_version` are allowed.
    extension: bool,
}

/// One admitted effective `BrtBeginExtConn14` binding.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct ExtConn14Binding {
    record: RecordSpan,
    wrapper: FrtWrapperSpan,
    culture_span: WideStringSpan,
    client_cube_urn: String,
    client_cube_urn_span: WideStringSpan,
}

impl ExtConn14Binding {
    /// Decoded Custom Data UID.  Empty values are legal and contribute no
    /// graph edge at the host layer.
    #[must_use]
    pub(crate) fn client_cube_urn(&self) -> &str {
        &self.client_cube_urn
    }

    /// Source range for `irstClientCubeUrn`.
    #[must_use]
    const fn client_cube_urn_span(&self) -> WideStringSpan {
        self.client_cube_urn_span
    }
}

/// A source-borrowing scan result.
#[derive(Debug, Clone)]
pub(crate) struct BindingScan<'a> {
    source: &'a [u8],
    records: usize,
    max_record_payload_bytes: usize,
    max_wrapper_depth: usize,
    max_calculated_member_records: usize,
    max_culture_units: usize,
    max_client_cube_urn_units: usize,
    work_bytes: usize,
    bindings: Vec<ExtConn14Binding>,
    opaque_blocks: Vec<OpaqueBlock>,
    blocking_opaque: bool,
}

impl<'a> BindingScan<'a> {
    /// Number of framed records scanned, including opaque records.
    #[must_use]
    pub(crate) const fn record_count(&self) -> usize {
        self.records
    }

    /// Borrow all effective ExtConn14 bindings in source order.
    #[must_use]
    pub(crate) fn bindings(&self) -> &[ExtConn14Binding] {
        &self.bindings
    }

    /// Whether a hidden or unsupported reference blocks identity mutation.
    #[must_use]
    pub(crate) const fn has_blocking_opaque(&self) -> bool {
        self.blocking_opaque
    }

    fn source_matches(&self, source: &[u8]) -> bool {
        source.as_ptr() == self.source.as_ptr() && source.len() == self.source.len()
    }
}

/// Caller-controlled resource ceilings for one binding scan and rewrite.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct BindingLimits {
    /// Maximum source-part bytes accepted by the scanner.
    pub(crate) max_source_bytes: usize,
    /// Maximum payload bytes in one BIFF12 record.
    pub(crate) max_record_payload_bytes: usize,
    /// Maximum framed records in one part.
    pub(crate) max_records: usize,
    /// Maximum effective ExtConn14 references.
    pub(crate) max_bindings: usize,
    /// Maximum opaque provenance blocks retained.
    pub(crate) max_opaque_blocks: usize,
    /// Maximum nested AC/FRT wrapper depth.
    pub(crate) max_wrapper_depth: usize,
    /// Maximum records in one ExtConn14 calculated-member collection.
    pub(crate) max_calculated_member_records: usize,
    /// Caller UTF-16 cap for culture strings; hard-capped at 84 units.
    pub(crate) max_culture_units: usize,
    /// Caller UTF-16 cap for Custom Data UIDs; hard-capped at 65,535 units.
    pub(crate) max_client_cube_urn_units: usize,
    /// Maximum candidate output bytes.
    pub(crate) max_output_bytes: usize,
    /// Maximum temporary planning/work bytes charged by this module.
    pub(crate) max_work_bytes: usize,
    /// Initial binding-vector reservation.  It is always clipped to the
    /// corresponding hard maximum.
    pub(crate) initial_binding_capacity: usize,
    /// Initial opaque-vector reservation.
    pub(crate) initial_opaque_capacity: usize,
}

impl BindingLimits {
    /// Validate caller ceilings against the wire hard limits.
    const fn validate(self) -> Result<Self> {
        if self.max_record_payload_bytes > MAX_WIRE_PAYLOAD {
            return Err(BindingError::InvalidLimit {
                name: "max_record_payload_bytes",
                value: self.max_record_payload_bytes,
                maximum: MAX_WIRE_PAYLOAD,
            });
        }
        if self.max_culture_units > HARD_CULTURE_UNITS {
            return Err(BindingError::InvalidLimit {
                name: "max_culture_units",
                value: self.max_culture_units,
                maximum: HARD_CULTURE_UNITS,
            });
        }
        if self.max_client_cube_urn_units > HARD_CLIENT_CUBE_URN_UNITS {
            return Err(BindingError::InvalidLimit {
                name: "max_client_cube_urn_units",
                value: self.max_client_cube_urn_units,
                maximum: HARD_CLIENT_CUBE_URN_UNITS,
            });
        }
        if self.max_bindings == 0 || self.max_records == 0 || self.max_wrapper_depth == 0 {
            return Err(BindingError::InvalidLimit {
                name: "nonzero scan ceilings",
                value: 0,
                maximum: usize::MAX,
            });
        }
        Ok(self)
    }
}

/// A caller-owned cancellation and work-accounting hook.
pub(crate) trait WorkBudget {
    /// Check cancellation before the next bounded unit of work.
    fn check(&mut self) -> Result<()> {
        Ok(())
    }

    /// Charge bounded work.  Implementations may reject before allocation.
    fn charge(&mut self, units: usize) -> Result<()> {
        self.check()?;
        let _ = units;
        Ok(())
    }
}

/// Scan a complete connections part with explicit PivotCache associations and
/// a caller-owned cancellation/work budget.
pub(crate) fn scan_connections_with_budget<'a, B: WorkBudget + ?Sized>(
    source: &'a [u8],
    limits: BindingLimits,
    pivot_cache_connection_ids: &[u32],
    budget: &mut B,
) -> Result<BindingScan<'a>> {
    let limits = limits.validate()?;
    if source.len() > limits.max_source_bytes {
        return Err(BindingError::SourceLimit {
            actual: source.len(),
            limit: limits.max_source_bytes,
        });
    }
    budget.check()?;

    let mut work = 0usize;
    let binding_capacity = limits
        .initial_binding_capacity
        .min(limits.max_bindings)
        .min(source.len().saturating_div(8));
    let opaque_capacity = limits
        .initial_opaque_capacity
        .min(limits.max_opaque_blocks)
        .min(source.len().saturating_div(4));
    let binding_reservation = binding_capacity
        .checked_mul(size_of::<ExtConn14Binding>())
        .ok_or(BindingError::LengthOverflow)?;
    let opaque_reservation = opaque_capacity
        .checked_mul(size_of::<OpaqueBlock>())
        .ok_or(BindingError::LengthOverflow)?;
    let stack_item_bytes = size_of::<FrtState>()
        .checked_add(size_of::<AcState>())
        .ok_or(BindingError::LengthOverflow)?;
    let stack_reservation = limits
        .max_wrapper_depth
        .checked_mul(stack_item_bytes)
        .ok_or(BindingError::LengthOverflow)?;
    charge_scan_work(
        &mut work,
        binding_reservation
            .checked_add(opaque_reservation)
            .and_then(|value| value.checked_add(stack_reservation))
            .ok_or(BindingError::LengthOverflow)?,
        &limits,
        budget,
    )?;
    let mut bindings = Vec::new();
    try_reserve(&mut bindings, binding_capacity, "ExtConn14 bindings")?;
    let mut opaque_blocks = Vec::new();
    try_reserve(
        &mut opaque_blocks,
        opaque_capacity,
        "opaque wrapper provenance",
    )?;
    let mut frt_stack: Vec<FrtState> = Vec::new();
    try_reserve(
        &mut frt_stack,
        limits.max_wrapper_depth,
        "future-record wrapper stack",
    )?;
    let mut ac_stack: Vec<AcState> = Vec::new();
    try_reserve(
        &mut ac_stack,
        limits.max_wrapper_depth,
        "alternate-content wrapper stack",
    )?;
    let mut cursor = 0usize;
    let mut records = 0usize;
    let first = next_record(source, &mut cursor, &limits)?;
    charge_scan_work(&mut work, first.span.range.len(), &limits, budget)?;
    records = records.checked_add(1).ok_or(BindingError::WorkLimit {
        actual: usize::MAX,
        limit: limits.max_records,
    })?;
    if first.span.kind != BEGIN_EXT_CONNECTIONS {
        return Err(BindingError::UnexpectedRecord {
            offset: first.span.range.start(),
            expected: BEGIN_EXT_CONNECTIONS,
            found: first.span.kind,
        });
    }
    require_empty_payload(first, "BrtBeginExtConnections")?;
    let mut root = ConnectionsRootSpan {
        begin: first.span,
        end: first.span,
        range: first.span.range,
    };

    let mut connection: Option<ConnectionState> = None;
    let mut db_props_active = false;
    let mut ext14: Option<Ext14State> = None;
    let mut ext15_active = false;
    let mut ended_root = false;
    let mut connection_ordinal = 0usize;
    let mut max_record_payload_bytes = first.span.payload.len();
    let mut max_wrapper_depth = 0usize;
    let mut max_calculated_member_records = 0usize;
    let mut max_culture_units = 0usize;
    let mut max_client_cube_urn_units = 0usize;

    while cursor < source.len() {
        budget.check()?;
        let record = next_record(source, &mut cursor, &limits)?;
        records = records.checked_add(1).ok_or(BindingError::WorkLimit {
            actual: usize::MAX,
            limit: limits.max_records,
        })?;
        if records > limits.max_records {
            return Err(BindingError::RecordLimit {
                actual: records,
                limit: limits.max_records,
            });
        }
        charge_scan_work(&mut work, record.span.range.len(), &limits, budget)?;
        max_record_payload_bytes = max_record_payload_bytes.max(record.span.payload.len());

        if ended_root {
            return Err(BindingError::TrailingRecords {
                offset: record.span.range.start(),
            });
        }

        if ext14.is_none()
            && !matches!(
                record.span.kind,
                END_EC_DB_PROPS | FRT_BEGIN | BEGIN_EXT_CONN14
            )
        {
            if let Some(state) = connection.as_mut() {
                state.db_props_recent = false;
            }
        }

        // ExtConn14's collection consumes every record until its matching end
        // marker.  Its calculated-member records are intentionally opaque;
        // the scalar begin-record grammar is the only pinned writer profile.
        if let Some(state) = ext14.as_mut() {
            if record.span.kind == END_EXT_CONN14 {
                require_empty_payload(record, "BrtEndExtConn14")?;
                let state = ext14.take().ok_or(BindingError::StateCorrupt)?;
                finish_ext14(
                    state,
                    &connection,
                    &mut frt_stack,
                    ext15_active,
                    pivot_cache_connection_ids,
                    &limits,
                    record.span,
                    &mut work,
                    budget,
                    &mut opaque_blocks,
                )?;
                continue;
            }
            state.child_records =
                state
                    .child_records
                    .checked_add(1)
                    .ok_or(BindingError::CalculatedMemberLimit {
                        actual: usize::MAX,
                        limit: limits.max_calculated_member_records,
                    })?;
            if state.child_records > limits.max_calculated_member_records {
                return Err(BindingError::CalculatedMemberLimit {
                    actual: state.child_records,
                    limit: limits.max_calculated_member_records,
                });
            }
            max_calculated_member_records = max_calculated_member_records.max(state.child_records);
            if !ac_stack.is_empty() {
                state.hidden = true;
                if let Some(ac) = ac_stack.last_mut() {
                    ac.contains_candidate = true;
                }
            }
            if record.span.kind == BEGIN_EXT_CONN14 || record.span.kind == BEGIN_EXT_CONN15 {
                state.hidden = true;
            }
            // Wrapper records nested in the opaque OLAP collection still need
            // balanced tracking, so the normal wrapper branches below run.
        }

        match record.span.kind {
            AC_BEGIN => {
                if ac_stack.len().saturating_add(frt_stack.len()) >= limits.max_wrapper_depth {
                    return Err(BindingError::WrapperDepthLimit {
                        limit: limits.max_wrapper_depth,
                    });
                }
                let ac = parse_ac_begin(record, &limits, &mut work, budget)?;
                max_wrapper_depth = max_wrapper_depth.max(
                    ac_stack
                        .len()
                        .saturating_add(frt_stack.len())
                        .saturating_add(1),
                );
                if let Some(state) = ext14.as_mut() {
                    state.hidden = true;
                    state.ac_nested = true;
                }
                ac_stack.push(AcState {
                    begin: record.span,
                    product_versions: ac,
                    contains_candidate: false,
                });
            },
            AC_END => {
                require_empty_payload(record, "BrtACEnd")?;
                let ac = ac_stack.pop().ok_or(BindingError::UnbalancedWrapper {
                    kind: AC_END,
                    offset: record.span.range.start(),
                })?;
                let ac_range = Span::new(ac.begin.range.start(), record.span.range.end())?;
                push_opaque(
                    &mut opaque_blocks,
                    OpaqueBlock {
                        range: ac_range,
                        reason: OpaqueReason::AlternateContent,
                        contains_reference_candidate: ac.contains_candidate,
                        frt_profile: None,
                        alternate_content_versions: ac.product_versions,
                    },
                    &limits,
                    &mut work,
                    budget,
                )?;
                if ac.contains_candidate {
                    if let Some(parent) = ac_stack.last_mut() {
                        parent.contains_candidate = true;
                    }
                }
            },
            FRT_BEGIN => {
                if ac_stack.len().saturating_add(frt_stack.len()) >= limits.max_wrapper_depth {
                    return Err(BindingError::WrapperDepthLimit {
                        limit: limits.max_wrapper_depth,
                    });
                }
                let profile = parse_frt_begin(record);
                let nested = !frt_stack.is_empty();
                let hidden = !ac_stack.is_empty() || nested || profile.is_err();
                max_wrapper_depth = max_wrapper_depth.max(
                    ac_stack
                        .len()
                        .saturating_add(frt_stack.len())
                        .saturating_add(1),
                );
                if let Some(state) = ext14.as_mut() {
                    state.hidden = true;
                }
                frt_stack.push(FrtState {
                    begin: record.span,
                    profile: profile.ok(),
                    known: !hidden,
                    seen_ext14: false,
                    seen_ext15: false,
                    pending: None,
                    hidden_candidate: false,
                });
            },
            FRT_END => {
                require_empty_payload(record, "BrtFRTEnd")?;
                let mut frt = frt_stack.pop().ok_or(BindingError::UnbalancedWrapper {
                    kind: FRT_END,
                    offset: record.span.range.start(),
                })?;
                // A future-record wrapper may occur inside an opaque
                // ExtConn14/15 collection.  Its matching end is balanced
                // when another FRT remains below it; only an outer wrapper
                // closing while the collection is still open is malformed.
                if (ext14.is_some() || ext15_active) && frt_stack.is_empty() {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: BEGIN_EXT_CONN14,
                        offset: record.span.range.start(),
                    });
                }
                if frt.known {
                    let wrapper = frt_wrapper(frt.begin, record.span, frt.profile)?;
                    if let Some(mut binding) = frt.pending.take() {
                        binding.wrapper = wrapper;
                        push_binding(&mut bindings, binding, &limits, &mut work, budget)?;
                    }
                } else if !frt.known {
                    let range = Span::new(frt.begin.range.start(), record.span.range.end())?;
                    push_opaque(
                        &mut opaque_blocks,
                        OpaqueBlock {
                            range,
                            reason: OpaqueReason::FutureRecordWrapper,
                            contains_reference_candidate: frt.hidden_candidate
                                || frt.pending.is_some(),
                            frt_profile: frt.profile,
                            alternate_content_versions: Vec::new(),
                        },
                        &limits,
                        &mut work,
                        budget,
                    )?;
                    if let Some(parent) = frt_stack.last_mut() {
                        parent.hidden_candidate = true;
                    }
                }
            },
            BEGIN_EXT_CONNECTIONS => {
                return Err(BindingError::UnexpectedRecord {
                    offset: record.span.range.start(),
                    expected: END_EXT_CONNECTIONS,
                    found: record.span.kind,
                });
            },
            END_EXT_CONNECTIONS => {
                require_empty_payload(record, "BrtEndExtConnections")?;
                if connection.is_some()
                    || db_props_active
                    || ext14.is_some()
                    || ext15_active
                    || !frt_stack.is_empty()
                    || !ac_stack.is_empty()
                {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: END_EXT_CONNECTIONS,
                        offset: record.span.range.start(),
                    });
                }
                root.end = record.span;
                root.range = Span::new(root.begin.range.start(), record.span.range.end())?;
                ended_root = true;
            },
            BEGIN_EXT_CONNECTION => {
                if connection.is_some() || !ac_stack.is_empty() {
                    return Err(BindingError::UnexpectedRecord {
                        offset: record.span.range.start(),
                        expected: END_EXT_CONNECTION,
                        found: record.span.kind,
                    });
                }
                if let Some(frt) = frt_stack.last_mut() {
                    // `BrtBeginExtConnection` is not part of the admitted
                    // FRTEXTCONNECTIONS child grammar.  Keep the complete FRT
                    // envelope opaque while still balancing this nested
                    // connection so the scanner can retain its provenance.
                    frt.known = false;
                    frt.hidden_candidate = true;
                }
                let (source_type, connection_id) = if frt_stack.is_empty() {
                    parse_connection_header(record)?
                } else {
                    // The connection header is outside the admitted FRT
                    // grammar.  Balance the envelope without interpreting
                    // its payload; the enclosing FRT is already opaque.
                    (0, 0)
                };
                connection_ordinal =
                    connection_ordinal
                        .checked_add(1)
                        .ok_or(BindingError::ConnectionLimit {
                            actual: usize::MAX,
                            limit: limits.max_bindings,
                        })?;
                connection = Some(ConnectionState {
                    source_type,
                    connection_id,
                    db_command: None,
                    db_props_recent: false,
                });
            },
            END_EXT_CONNECTION => {
                require_empty_payload(record, "BrtEndExtConnection")?;
                let active_known_frt = frt_stack.last().is_some_and(|frt| frt.known);
                if ext14.is_some() || ext15_active || active_known_frt {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: END_EXT_CONNECTION,
                        offset: record.span.range.start(),
                    });
                }
                if connection.take().is_none() {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: END_EXT_CONNECTION,
                        offset: record.span.range.start(),
                    });
                }
                db_props_active = false;
            },
            BEGIN_EC_DB_PROPS => {
                if connection.is_none() || db_props_active {
                    return Err(BindingError::Context {
                        offset: record.span.range.start(),
                        detail: "BrtBeginECDbProps is outside one connection",
                    });
                }
                let command_type = read_u32(record.payload_bytes, 0)?;
                db_props_active = true;
                if let Some(state) = connection.as_mut() {
                    state.db_command = Some(command_type);
                }
            },
            END_EC_DB_PROPS => {
                require_empty_payload(record, "BrtEndECDbProps")?;
                if !db_props_active {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: END_EC_DB_PROPS,
                        offset: record.span.range.start(),
                    });
                }
                db_props_active = false;
                if let Some(state) = connection.as_mut() {
                    state.db_props_recent = true;
                }
            },
            BEGIN_EXT_CONN14 => {
                begin_ext14(
                    record,
                    &connection,
                    &mut frt_stack,
                    &mut ac_stack,
                    ext15_active,
                    &limits,
                    &mut max_culture_units,
                    &mut max_client_cube_urn_units,
                    &mut work,
                    budget,
                    &mut ext14,
                )?;
            },
            END_EXT_CONN14 => {
                return Err(BindingError::UnbalancedWrapper {
                    kind: END_EXT_CONN14,
                    offset: record.span.range.start(),
                });
            },
            BEGIN_EXT_CONN15 => {
                if let Some(state) = ext14.as_mut() {
                    state.hidden = true;
                    state.ext15_nested = true;
                }
                begin_ext15(record, &mut frt_stack, &ac_stack, &mut ext15_active)?;
            },
            END_EXT_CONN15 => {
                require_empty_payload(record, "BrtEndExtConn15")?;
                if !ext15_active {
                    return Err(BindingError::UnbalancedWrapper {
                        kind: END_EXT_CONN15,
                        offset: record.span.range.start(),
                    });
                }
                ext15_active = false;
            },
            _ => {
                if ext14.is_none() && !ext15_active {
                    if let Some(state) = frt_stack.last_mut() {
                        if state.known
                            && !matches!(
                                record.span.kind,
                                END_EXT_CONN14 | END_EXT_CONN15 | FRT_END
                            )
                        {
                            // The pinned wrapper admits only ExtConn14/15 between
                            // its begin/end records.  An unrelated future record
                            // makes the wrapper opaque, but it remains readable.
                            state.known = false;
                        }
                    }
                }
            },
        }

        if let Some(state) = frt_stack.last_mut() {
            if !ac_stack.is_empty() {
                state.known = false;
            }
        }
    }

    if !ended_root {
        return Err(BindingError::UnexpectedEnd {
            context: "BrtEndExtConnections",
        });
    }
    if !opaque_blocks.is_empty() {
        // Keep this check after all spans are known, since wrapper blocks are
        // appended only when their matching end records are seen.
        validate_opaque_spans(&opaque_blocks, source.len())?;
    }
    validate_binding_spans(&bindings, source.len())?;
    let blocking_opaque = opaque_blocks_have_candidates(&opaque_blocks);
    Ok(BindingScan {
        source,
        records,
        max_record_payload_bytes,
        max_wrapper_depth,
        max_calculated_member_records,
        max_culture_units,
        max_client_cube_urn_units,
        work_bytes: work,
        bindings,
        opaque_blocks,
        blocking_opaque,
    })
}

/// One source-to-source UID replacement requested by a host transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct UidRewrite<'a> {
    /// Existing decoded UID to match.
    pub(crate) from: &'a str,
    /// Replacement decoded UID.  Empty is valid for detach operations.
    pub(crate) to: &'a str,
}

/// Result of one ordered UID splice.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct RewriteOutcome {
    /// Whether at least one field's bytes changed.
    pub(crate) changed: bool,
    /// Number of ExtConn14 fields rewritten.
    pub(crate) replacements: usize,
    /// Final output length.
    pub(crate) output_len: usize,
}

/// Rewrite all effective references selected by a bounded list of UID edits
/// under a caller-owned cancellation and work budget.
///
/// The edit index borrows every `from`/`to` string.  Its open-addressed table
/// has a bounded load factor, so duplicate detection and binding lookup are
/// linear in the edit and binding counts (plus the bytes actually compared).
/// All budget and output-limit checks happen before `output` is cleared or
/// any source bytes are copied.
pub(crate) fn rewrite_client_cube_urns_with_budget<B: WorkBudget + ?Sized>(
    source: &[u8],
    scan: &BindingScan<'_>,
    edits: &[UidRewrite<'_>],
    output: &mut Vec<u8>,
    limits: BindingLimits,
    budget: &mut B,
) -> Result<RewriteOutcome> {
    let limits = limits.validate()?;
    if source.len() > limits.max_source_bytes {
        return Err(BindingError::SourceLimit {
            actual: source.len(),
            limit: limits.max_source_bytes,
        });
    }
    if !scan.source_matches(source) {
        return Err(BindingError::SourceMismatch);
    }
    if scan.bindings.len() > limits.max_bindings {
        return Err(BindingError::BindingLimit {
            actual: scan.bindings.len(),
            limit: limits.max_bindings,
        });
    }
    if scan.records > limits.max_records {
        return Err(BindingError::RecordLimit {
            actual: scan.records,
            limit: limits.max_records,
        });
    }
    if scan.max_record_payload_bytes > limits.max_record_payload_bytes {
        return Err(BindingError::RecordPayloadLimit {
            actual: scan.max_record_payload_bytes,
            limit: limits.max_record_payload_bytes,
        });
    }
    if scan.max_wrapper_depth > limits.max_wrapper_depth {
        return Err(BindingError::WrapperDepthLimit {
            limit: limits.max_wrapper_depth,
        });
    }
    if scan.max_calculated_member_records > limits.max_calculated_member_records {
        return Err(BindingError::CalculatedMemberLimit {
            actual: scan.max_calculated_member_records,
            limit: limits.max_calculated_member_records,
        });
    }
    if scan.max_culture_units > limits.max_culture_units {
        return Err(BindingError::StringLimit {
            field: "irstCulture",
            units: scan.max_culture_units,
            limit: limits.max_culture_units.min(HARD_CULTURE_UNITS),
            offset: 0,
        });
    }
    if scan.max_client_cube_urn_units > limits.max_client_cube_urn_units {
        return Err(BindingError::StringLimit {
            field: "irstClientCubeUrn",
            units: scan.max_client_cube_urn_units,
            limit: limits
                .max_client_cube_urn_units
                .min(HARD_CLIENT_CUBE_URN_UNITS),
            offset: 0,
        });
    }
    if scan.work_bytes > limits.max_work_bytes {
        return Err(BindingError::WorkLimit {
            actual: scan.work_bytes,
            limit: limits.max_work_bytes,
        });
    }
    if scan.opaque_blocks.len() > limits.max_opaque_blocks {
        return Err(BindingError::OpaqueLimit {
            actual: scan.opaque_blocks.len(),
            limit: limits.max_opaque_blocks,
        });
    }
    if scan.has_blocking_opaque() {
        return Err(BindingError::HiddenReference);
    }
    if edits.len() > limits.max_bindings {
        return Err(BindingError::RewriteCountLimit {
            actual: edits.len(),
            limit: limits.max_bindings,
        });
    }
    if edits.is_empty() {
        return Ok(RewriteOutcome {
            changed: false,
            replacements: 0,
            output_len: source.len(),
        });
    }

    budget.check()?;
    // The scan's bounded source walk and retained metadata are part of the
    // same operation budget.  They were already charged to `budget`; carry
    // their measured amount into the rewrite ceiling without charging it a
    // second time.
    let mut work = scan.work_bytes;
    let index = BorrowedUidIndex::build(edits, &mut work, &limits, budget)?;

    let mut match_count = 0usize;
    let mut final_len = source.len();
    for binding in &scan.bindings {
        let Some(edit_index) =
            index.find(binding.client_cube_urn(), edits, &mut work, &limits, budget)?
        else {
            continue;
        };
        let edit = &edits[edit_index];
        if charged_str_equal(
            edit.to,
            binding.client_cube_urn(),
            &mut work,
            &limits,
            budget,
        )? {
            continue;
        }
        let old_wire = binding.client_cube_urn_span().field.len();
        let target_units = index.target_units(edit_index)?;
        let new_wire = wire_wide_string_len_from_units(target_units)?;
        let payload_len = binding
            .record
            .payload
            .len()
            .checked_sub(old_wire)
            .and_then(|value| value.checked_add(new_wire))
            .ok_or(BindingError::LengthOverflow)?;
        if payload_len > limits.max_record_payload_bytes || payload_len > MAX_WIRE_PAYLOAD {
            return Err(BindingError::RecordPayloadLimit {
                actual: payload_len,
                limit: limits.max_record_payload_bytes,
            });
        }
        let length_len = if payload_len == binding.record.payload.len() {
            usize::from(binding.record.length_len)
        } else {
            encoded_varint_len(payload_len)
        };
        let new_record_len = usize::from(binding.record.kind_len)
            .checked_add(length_len)
            .and_then(|value| value.checked_add(payload_len))
            .ok_or(BindingError::LengthOverflow)?;
        let old_record_len = binding.record.range.len();
        if new_record_len >= old_record_len {
            final_len = final_len
                .checked_add(new_record_len - old_record_len)
                .ok_or(BindingError::LengthOverflow)?;
        } else {
            final_len = final_len
                .checked_sub(old_record_len - new_record_len)
                .ok_or(BindingError::LengthOverflow)?;
        }
        match_count = match_count
            .checked_add(1)
            .ok_or(BindingError::RewriteCountLimit {
                actual: usize::MAX,
                limit: limits.max_bindings,
            })?;
        charge_rewrite_work(
            &mut work,
            size_of::<PlannedReplacement<'_>>(),
            &limits,
            budget,
        )?;
    }
    if match_count == 0 {
        return Ok(RewriteOutcome {
            changed: false,
            replacements: 0,
            output_len: source.len(),
        });
    }
    if final_len > limits.max_output_bytes {
        return Err(BindingError::OutputLimit {
            actual: final_len,
            limit: limits.max_output_bytes,
        });
    }

    // The final source read and output write are charged before output
    // allocation/mutation.  The UTF-16 wire loops are charged below after
    // their ordered replacement plan is built, still before output mutation.
    // This keeps cancellation and work-limit refusal atomic while accounting
    // for the actual bytes and units used.
    charge_rewrite_work(&mut work, source.len(), &limits, budget)?;
    charge_rewrite_work(&mut work, final_len, &limits, budget)?;

    let mut replacements = Vec::new();
    try_reserve(&mut replacements, match_count, "ExtConn14 rewrite plan")?;
    for binding in &scan.bindings {
        let Some(edit_index) =
            index.find(binding.client_cube_urn(), edits, &mut work, &limits, budget)?
        else {
            continue;
        };
        let edit = &edits[edit_index];
        if charged_str_equal(
            edit.to,
            binding.client_cube_urn(),
            &mut work,
            &limits,
            budget,
        )? {
            continue;
        }
        let units = index.target_units(edit_index)?;
        let replacement_wire_len = wire_wide_string_len_from_units(units)?;
        let payload_len = binding
            .record
            .payload
            .len()
            .checked_sub(binding.client_cube_urn_span().field.len())
            .and_then(|value| value.checked_add(replacement_wire_len))
            .ok_or(BindingError::LengthOverflow)?;
        replacements.push(PlannedReplacement {
            record: binding.record,
            string: binding.client_cube_urn_span(),
            target: edit.to,
            target_units: units,
            payload_len,
        });
    }
    validate_replacement_order(&replacements)?;

    // Count the exact UTF-16 wire iterations that the splice will perform.
    // Keeping this preflight separate from the write loop means a budget can
    // reject midway through a large many-to-one batch without changing the
    // caller's output buffer.
    for replacement in &replacements {
        let mut units = 0usize;
        for _unit in replacement.target.encode_utf16() {
            units = units.checked_add(1).ok_or(BindingError::LengthOverflow)?;
            charge_rewrite_work(&mut work, 1, &limits, budget)?;
        }
        if units != replacement.target_units {
            return Err(BindingError::StateCorrupt);
        }
    }

    // The planning pass above is deliberately repeated to retain source
    // order in the splice.  It performs only bounded borrowed-index lookups;
    // charge its byte comparisons before the first output mutation as well.
    // The source/output and UTF-16 charges above cover the actual copy/encode
    // loops, while these lookup charges cover this second pass.
    if output.capacity() < final_len {
        output
            .try_reserve_exact(final_len - output.capacity())
            .map_err(|_error| BindingError::Allocation {
                resource: "rewritten connections output",
            })?;
    }
    output.clear();
    let mut source_cursor = 0usize;
    for replacement in &replacements {
        output.extend_from_slice(
            source
                .get(source_cursor..replacement.record.range.start())
                .ok_or(BindingError::SourceMismatch)?,
        );
        let kind_end = replacement
            .record
            .range
            .start()
            .checked_add(usize::from(replacement.record.kind_len))
            .ok_or(BindingError::LengthOverflow)?;
        output.extend_from_slice(
            source
                .get(replacement.record.range.start()..kind_end)
                .ok_or(BindingError::SourceMismatch)?,
        );
        if replacement.payload_len == replacement.record.payload.len() {
            output.extend_from_slice(
                source
                    .get(kind_end..replacement.record.payload.start())
                    .ok_or(BindingError::SourceMismatch)?,
            );
        } else {
            encode_varint(replacement.payload_len, output)?;
        }
        output.extend_from_slice(
            source
                .get(replacement.record.payload.start()..replacement.string.field.start())
                .ok_or(BindingError::SourceMismatch)?,
        );
        let units = u32::try_from(replacement.target_units)
            .map_err(|_error| BindingError::LengthOverflow)?;
        output.extend_from_slice(&units.to_le_bytes());
        for unit in replacement.target.encode_utf16() {
            output.extend_from_slice(&unit.to_le_bytes());
        }
        output.extend_from_slice(
            source
                .get(replacement.string.field.end()..replacement.record.range.end())
                .ok_or(BindingError::SourceMismatch)?,
        );
        source_cursor = replacement.record.range.end();
    }
    output.extend_from_slice(
        source
            .get(source_cursor..)
            .ok_or(BindingError::SourceMismatch)?,
    );
    debug_assert_eq!(output.len(), final_len);
    Ok(RewriteOutcome {
        changed: true,
        replacements: replacements.len(),
        output_len: output.len(),
    })
}

#[derive(Debug, Clone, Copy)]
struct RawRecord<'a> {
    span: RecordSpan,
    payload_bytes: &'a [u8],
}

#[derive(Debug)]
struct ConnectionState {
    source_type: u32,
    connection_id: u32,
    db_command: Option<u32>,
    db_props_recent: bool,
}

#[derive(Debug)]
struct Ext14State {
    record: RecordSpan,
    parsed: Option<ParsedExt14>,
    child_records: usize,
    hidden: bool,
    ac_nested: bool,
    ext15_nested: bool,
    db_props_recent: bool,
}

#[derive(Debug)]
struct ParsedExt14 {
    culture_span: WideStringSpan,
    client_cube_urn: String,
    client_cube_urn_span: WideStringSpan,
}

#[derive(Debug)]
struct FrtState {
    begin: RecordSpan,
    profile: Option<FrtProductVersion>,
    known: bool,
    seen_ext14: bool,
    seen_ext15: bool,
    pending: Option<ExtConn14Binding>,
    hidden_candidate: bool,
}

#[derive(Debug)]
struct AcState {
    begin: RecordSpan,
    product_versions: Vec<AcProductVersion>,
    contains_candidate: bool,
}

#[derive(Debug, Clone, Copy)]
struct PlannedReplacement<'a> {
    record: RecordSpan,
    string: WideStringSpan,
    target: &'a str,
    target_units: usize,
    payload_len: usize,
}

#[derive(Debug, Clone, Copy)]
struct IndexedUid {
    hash: usize,
    edit: usize,
}

/// A borrowed, bounded UID lookup table for one rewrite batch.
///
/// The table owns only slot metadata and UTF-16 unit counts.  Both UID
/// strings remain borrowed from the caller's edit slice, so building the
/// index cannot copy a source or edit payload.
#[derive(Debug)]
struct BorrowedUidIndex {
    slots: Vec<Option<IndexedUid>>,
    target_units: Vec<usize>,
    mask: usize,
}

impl BorrowedUidIndex {
    fn build<B: WorkBudget + ?Sized>(
        edits: &[UidRewrite<'_>],
        work: &mut usize,
        limits: &BindingLimits,
        budget: &mut B,
    ) -> Result<Self> {
        let mut target_units = Vec::new();
        let unit_reservation = edits
            .len()
            .checked_mul(size_of::<usize>())
            .ok_or(BindingError::LengthOverflow)?;
        charge_rewrite_work(work, unit_reservation, limits, budget)?;
        target_units
            .try_reserve_exact(edits.len())
            .map_err(|_error| BindingError::Allocation {
                resource: "ExtConn14 rewrite UID metadata",
            })?;
        for edit in edits {
            let units = count_utf16_units(edit.to, work, limits, budget)?;
            if units > limits.max_client_cube_urn_units {
                return Err(BindingError::UidLimit {
                    units,
                    limit: limits.max_client_cube_urn_units,
                });
            }
            target_units.push(units);
        }

        let capacity = uid_index_capacity(edits.len())?;
        let slot_reservation = capacity
            .checked_mul(size_of::<Option<IndexedUid>>())
            .ok_or(BindingError::LengthOverflow)?;
        charge_rewrite_work(work, slot_reservation, limits, budget)?;
        let mut slots = Vec::new();
        slots
            .try_reserve_exact(capacity)
            .map_err(|_error| BindingError::Allocation {
                resource: "ExtConn14 rewrite UID index",
            })?;
        slots.resize(capacity, None);
        let mask = capacity - 1;

        for (edit_index, edit) in edits.iter().enumerate() {
            let hash = hash_uid(edit.from, work, limits, budget)?;
            let mut slot = hash & mask;
            let mut inserted = false;
            for _ in 0..capacity {
                charge_rewrite_work(work, 1, limits, budget)?;
                match slots[slot] {
                    None => {
                        slots[slot] = Some(IndexedUid {
                            hash,
                            edit: edit_index,
                        });
                        inserted = true;
                        break;
                    },
                    Some(existing) => {
                        if existing.hash == hash
                            && charged_str_equal(
                                edit.from,
                                edits[existing.edit].from,
                                work,
                                limits,
                                budget,
                            )?
                        {
                            if charged_str_equal(
                                edit.to,
                                edits[existing.edit].to,
                                work,
                                limits,
                                budget,
                            )? {
                                return Err(BindingError::DuplicateRewrite);
                            }
                            return Err(BindingError::ConflictingRewrite);
                        }
                        slot = (slot + 1) & mask;
                    },
                }
            }
            if !inserted {
                return Err(BindingError::StateCorrupt);
            }
        }
        Ok(Self {
            slots,
            target_units,
            mask,
        })
    }

    fn target_units(&self, edit: usize) -> Result<usize> {
        self.target_units
            .get(edit)
            .copied()
            .ok_or(BindingError::StateCorrupt)
    }

    fn find<B: WorkBudget + ?Sized>(
        &self,
        uid: &str,
        edits: &[UidRewrite<'_>],
        work: &mut usize,
        limits: &BindingLimits,
        budget: &mut B,
    ) -> Result<Option<usize>> {
        let hash = hash_uid(uid, work, limits, budget)?;
        let mut slot = hash & self.mask;
        for _ in 0..self.slots.len() {
            charge_rewrite_work(work, 1, limits, budget)?;
            let Some(existing) = self.slots[slot] else {
                return Ok(None);
            };
            if existing.hash == hash
                && charged_str_equal(uid, edits[existing.edit].from, work, limits, budget)?
            {
                return Ok(Some(existing.edit));
            }
            slot = (slot + 1) & self.mask;
        }
        Ok(None)
    }
}

fn uid_index_capacity(edit_count: usize) -> Result<usize> {
    if edit_count == 0 {
        return Ok(1);
    }
    edit_count
        .checked_mul(2)
        .and_then(usize::checked_next_power_of_two)
        .ok_or(BindingError::LengthOverflow)
}

fn charge_rewrite_work<B: WorkBudget + ?Sized>(
    work: &mut usize,
    amount: usize,
    limits: &BindingLimits,
    budget: &mut B,
) -> Result<()> {
    let next = work
        .checked_add(amount)
        .ok_or(BindingError::LengthOverflow)?;
    if next > limits.max_work_bytes {
        return Err(BindingError::WorkLimit {
            actual: next,
            limit: limits.max_work_bytes,
        });
    }
    *work = next;
    budget.charge(amount)
}

fn charge_scan_work<B: WorkBudget + ?Sized>(
    work: &mut usize,
    amount: usize,
    limits: &BindingLimits,
    budget: &mut B,
) -> Result<()> {
    let next = work
        .checked_add(amount)
        .ok_or(BindingError::LengthOverflow)?;
    if next > limits.max_work_bytes {
        return Err(BindingError::WorkLimit {
            actual: next,
            limit: limits.max_work_bytes,
        });
    }
    *work = next;
    budget.charge(amount)
}

fn count_utf16_units<B: WorkBudget + ?Sized>(
    value: &str,
    work: &mut usize,
    limits: &BindingLimits,
    budget: &mut B,
) -> Result<usize> {
    let mut units = 0usize;
    for _unit in value.encode_utf16() {
        units = units.checked_add(1).ok_or(BindingError::LengthOverflow)?;
        charge_rewrite_work(work, 1, limits, budget)?;
    }
    Ok(units)
}

fn hash_uid<B: WorkBudget + ?Sized>(
    value: &str,
    work: &mut usize,
    limits: &BindingLimits,
    budget: &mut B,
) -> Result<usize> {
    // FNV-1a is deterministic and keeps collision handling explicit in the
    // bounded table.  The byte loop is charged one unit at a time so a
    // caller can cancel a pathological edit batch before output mutation.
    let mut hash = 0x811c_9dc5usize;
    for byte in value.as_bytes() {
        charge_rewrite_work(work, 1, limits, budget)?;
        hash ^= usize::from(*byte);
        hash = hash.wrapping_mul(0x0100_0193usize);
    }
    Ok(hash)
}

fn charged_str_equal<B: WorkBudget + ?Sized>(
    left: &str,
    right: &str,
    work: &mut usize,
    limits: &BindingLimits,
    budget: &mut B,
) -> Result<bool> {
    if left.len() != right.len() {
        return Ok(false);
    }
    for (left, right) in left.as_bytes().iter().zip(right.as_bytes()) {
        charge_rewrite_work(work, 1, limits, budget)?;
        if left != right {
            return Ok(false);
        }
    }
    Ok(true)
}

fn begin_ext14<'a, B: WorkBudget + ?Sized>(
    record: RawRecord<'a>,
    connection: &Option<ConnectionState>,
    frt_stack: &mut [FrtState],
    ac_stack: &mut [AcState],
    ext15_active: bool,
    limits: &BindingLimits,
    max_culture_units: &mut usize,
    max_client_cube_urn_units: &mut usize,
    work: &mut usize,
    budget: &mut B,
    output: &mut Option<Ext14State>,
) -> Result<()> {
    if output.is_some() {
        return Err(BindingError::UnbalancedWrapper {
            kind: BEGIN_EXT_CONN14,
            offset: record.span.range.start(),
        });
    }
    let direct_depth = frt_stack.len();
    let Some(frt) = frt_stack.last_mut() else {
        if let Some(ac) = ac_stack.last_mut() {
            ac.contains_candidate = true;
        }
        *output = Some(Ext14State {
            record: record.span,
            parsed: None,
            child_records: 0,
            hidden: true,
            ac_nested: !ac_stack.is_empty(),
            ext15_nested: ext15_active,
            db_props_recent: false,
        });
        return Ok(());
    };
    let direct_known = direct_depth == 1
        && frt.known
        && ac_stack.is_empty()
        && !ext15_active
        && !frt.seen_ext14
        && !frt.seen_ext15
        && connection.is_some();
    if frt.seen_ext14 || frt.seen_ext15 {
        frt.known = false;
    }
    frt.seen_ext14 = true;
    if direct_known {
        let parsed = parse_ext14(record, limits, work, budget)?;
        *max_culture_units = (*max_culture_units).max(parsed.culture_span.units);
        *max_client_cube_urn_units =
            (*max_client_cube_urn_units).max(parsed.client_cube_urn_span.units);
        *output = Some(Ext14State {
            record: record.span,
            parsed: Some(parsed),
            child_records: 0,
            hidden: false,
            ac_nested: false,
            ext15_nested: false,
            db_props_recent: connection
                .as_ref()
                .is_some_and(|state| state.db_props_recent),
        });
    } else {
        frt.known = false;
        frt.hidden_candidate = true;
        *output = Some(Ext14State {
            record: record.span,
            parsed: None,
            child_records: 0,
            hidden: true,
            ac_nested: !ac_stack.is_empty(),
            ext15_nested: ext15_active,
            db_props_recent: false,
        });
    }
    Ok(())
}

fn begin_ext15<'a>(
    record: RawRecord<'a>,
    frt_stack: &mut [FrtState],
    ac_stack: &[AcState],
    ext15_active: &mut bool,
) -> Result<()> {
    if *ext15_active {
        return Err(BindingError::UnbalancedWrapper {
            kind: BEGIN_EXT_CONN15,
            offset: record.span.range.start(),
        });
    }
    let direct_depth = frt_stack.len();
    let Some(frt) = frt_stack.last_mut() else {
        *ext15_active = true;
        return Ok(());
    };
    if direct_depth != 1 || !frt.known || !ac_stack.is_empty() || frt.seen_ext15 {
        frt.known = false;
    }
    frt.seen_ext15 = true;
    *ext15_active = true;
    Ok(())
}

fn finish_ext14<B: WorkBudget + ?Sized>(
    state: Ext14State,
    connection: &Option<ConnectionState>,
    frt_stack: &mut [FrtState],
    ext15_active: bool,
    pivot_cache_connection_ids: &[u32],
    limits: &BindingLimits,
    end: RecordSpan,
    work: &mut usize,
    budget: &mut B,
    opaque_blocks: &mut Vec<OpaqueBlock>,
) -> Result<()> {
    let opaque_range = Span::new(state.record.range.start(), end.range.end())?;
    if frt_stack.is_empty() {
        push_opaque(
            opaque_blocks,
            OpaqueBlock {
                range: opaque_range,
                reason: if state.ext15_nested {
                    OpaqueReason::ExtConn15
                } else if state.ac_nested {
                    OpaqueReason::AlternateContent
                } else {
                    OpaqueReason::UnsupportedExtConn14
                },
                contains_reference_candidate: true,
                frt_profile: None,
                alternate_content_versions: Vec::new(),
            },
            limits,
            work,
            budget,
        )?;
        return Ok(());
    }
    let Some(frt) = frt_stack.last() else {
        return Err(BindingError::StateCorrupt);
    };
    let Some(connection) = connection.as_ref() else {
        push_opaque(
            opaque_blocks,
            OpaqueBlock {
                range: opaque_range,
                reason: OpaqueReason::Context,
                contains_reference_candidate: true,
                frt_profile: frt.profile,
                alternate_content_versions: Vec::new(),
            },
            limits,
            work,
            budget,
        )?;
        return Ok(());
    };
    let pivot_cache = pivot_cache_connection_ids.contains(&connection.connection_id);
    let context_ok = !state.hidden
        && !ext15_active
        && frt_stack.len() == 1
        && frt.known
        && connection.source_type == DBTOLEDB
        && state.parsed.is_some()
        && (state.child_records == 0
            || (connection.db_command == Some(CMDCUBE) && state.db_props_recent))
        && (!pivot_cache || state.child_records == 0);
    let Some(parsed) = state.parsed else {
        push_opaque(
            opaque_blocks,
            OpaqueBlock {
                range: opaque_range,
                reason: OpaqueReason::UnsupportedExtConn14,
                contains_reference_candidate: true,
                frt_profile: frt.profile,
                alternate_content_versions: Vec::new(),
            },
            limits,
            work,
            budget,
        )?;
        return Ok(());
    };
    if !context_ok {
        let reason = if state.ext15_nested {
            OpaqueReason::ExtConn15
        } else if pivot_cache && state.child_records != 0 {
            OpaqueReason::Context
        } else if state.ac_nested {
            OpaqueReason::AlternateContent
        } else if !frt.known || frt_stack.len() != 1 {
            OpaqueReason::FutureRecordWrapper
        } else {
            OpaqueReason::Context
        };
        push_opaque(
            opaque_blocks,
            OpaqueBlock {
                range: opaque_range,
                reason,
                contains_reference_candidate: true,
                frt_profile: frt.profile,
                alternate_content_versions: Vec::new(),
            },
            limits,
            work,
            budget,
        )?;
        return Ok(());
    }
    let binding = ExtConn14Binding {
        record: state.record,
        wrapper: FrtWrapperSpan {
            begin: frt.begin,
            end: frt.begin,
            range: frt.begin.range,
            profile: frt.profile.ok_or(BindingError::StateCorrupt)?,
        },
        culture_span: parsed.culture_span,
        client_cube_urn: parsed.client_cube_urn,
        client_cube_urn_span: parsed.client_cube_urn_span,
    };
    if let Some(frt) = frt_stack.last_mut() {
        if frt.pending.is_some() {
            return Err(BindingError::DuplicateBindingSpan);
        }
        frt.pending = Some(binding);
    }
    Ok(())
}

fn parse_ext14<'a, B: WorkBudget + ?Sized>(
    record: RawRecord<'a>,
    limits: &BindingLimits,
    work: &mut usize,
    budget: &mut B,
) -> Result<ParsedExt14> {
    let payload = record.payload_bytes;
    if payload.len() < 4 {
        return Err(BindingError::Truncated {
            offset: record.span.payload.start(),
            needed: 4,
            available: payload.len(),
            context: "BrtBeginExtConn14 FRTBlank",
        });
    }
    // FRTBlank is a reserved field.  Its bytes belong to the admitted
    // grammar and are retained in the source splice, but the grammar says
    // they are ignored by consumers; do not reject a non-zero value here.
    let (_culture, culture_span, next) = parse_wide_string(
        payload,
        4,
        limits,
        limits.max_culture_units,
        HARD_CULTURE_UNITS,
        record.span.payload.start(),
        "irstCulture",
        work,
        budget,
    )?;
    let (client_cube_urn, client_cube_urn_span, end) = parse_wide_string(
        payload,
        next,
        limits,
        limits.max_client_cube_urn_units,
        HARD_CLIENT_CUBE_URN_UNITS,
        record.span.payload.start(),
        "irstClientCubeUrn",
        work,
        budget,
    )?;
    if end != payload.len() {
        return Err(BindingError::TrailingPayload {
            offset: record.span.payload.start() + end,
            remaining: payload.len() - end,
            context: "BrtBeginExtConn14",
        });
    }
    Ok(ParsedExt14 {
        culture_span,
        client_cube_urn,
        client_cube_urn_span,
    })
}

fn parse_wide_string<B: WorkBudget + ?Sized>(
    payload: &[u8],
    offset: usize,
    limits: &BindingLimits,
    caller_limit: usize,
    hard_limit: usize,
    absolute_base: usize,
    field: &'static str,
    work: &mut usize,
    budget: &mut B,
) -> Result<(String, WideStringSpan, usize)> {
    let units_u32 = read_u32(payload, offset)?;
    let units = usize::try_from(units_u32).map_err(|_error| BindingError::LengthOverflow)?;
    if units > caller_limit || units > hard_limit {
        return Err(BindingError::StringLimit {
            field,
            units,
            limit: caller_limit.min(hard_limit),
            offset: absolute_base + offset,
        });
    }
    let bytes = units.checked_mul(2).ok_or(BindingError::LengthOverflow)?;
    let data_start = offset.checked_add(4).ok_or(BindingError::LengthOverflow)?;
    let end = data_start
        .checked_add(bytes)
        .ok_or(BindingError::LengthOverflow)?;
    let absolute_offset = absolute_base
        .checked_add(offset)
        .ok_or(BindingError::LengthOverflow)?;
    let absolute_data_start = absolute_base
        .checked_add(data_start)
        .ok_or(BindingError::LengthOverflow)?;
    let absolute_end = absolute_base
        .checked_add(end)
        .ok_or(BindingError::LengthOverflow)?;
    let raw = payload
        .get(data_start..end)
        .ok_or(BindingError::Truncated {
            offset: absolute_data_start,
            needed: bytes,
            available: payload.len().saturating_sub(data_start),
            context: field,
        })?;
    let decoded_reservation = decoded_utf8_len(raw, absolute_data_start, field)?;
    charge_scan_work(work, decoded_reservation, limits, budget)?;
    let mut decoded = String::new();
    decoded
        .try_reserve_exact(decoded_reservation)
        .map_err(|_error| BindingError::Allocation { resource: field })?;
    let mut index = 0usize;
    while index < raw.len() {
        let unit = u16::from_le_bytes([raw[index], raw[index + 1]]);
        let value = if (0xd800..=0xdbff).contains(&unit) {
            let next_index = index.checked_add(2).ok_or(BindingError::LengthOverflow)?;
            let next = raw
                .get(next_index..next_index + 2)
                .ok_or(BindingError::InvalidUtf16 {
                    offset: absolute_data_start + index,
                })?;
            let low = u16::from_le_bytes([next[0], next[1]]);
            if !(0xdc00..=0xdfff).contains(&low) {
                return Err(BindingError::InvalidUtf16 {
                    offset: absolute_data_start + index,
                });
            }
            index = next_index + 2;
            let high = u32::from(unit) - 0xd800;
            let low = u32::from(low) - 0xdc00;
            char::from_u32(0x1_0000 + (high << 10) + low).ok_or(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index - 4,
            })?
        } else if (0xdc00..=0xdfff).contains(&unit) {
            return Err(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index,
            });
        } else {
            index = index.checked_add(2).ok_or(BindingError::LengthOverflow)?;
            char::from_u32(u32::from(unit)).ok_or(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index - 2,
            })?
        };
        decoded.push(value);
    }
    let span = WideStringSpan {
        field: Span::new(absolute_offset, absolute_end)?,
        length: Span::new(absolute_offset, absolute_data_start)?,
        data: Span::new(absolute_data_start, absolute_end)?,
        units,
    };
    Ok((decoded, span, end))
}

fn decoded_utf8_len(raw: &[u8], absolute_data_start: usize, field: &'static str) -> Result<usize> {
    let mut bytes = 0usize;
    let mut index = 0usize;
    while index < raw.len() {
        let unit = u16::from_le_bytes([raw[index], raw[index + 1]]);
        let value = if (0xd800..=0xdbff).contains(&unit) {
            let next_index = index.checked_add(2).ok_or(BindingError::LengthOverflow)?;
            let next = raw
                .get(next_index..next_index + 2)
                .ok_or(BindingError::InvalidUtf16 {
                    offset: absolute_data_start + index,
                })?;
            let low = u16::from_le_bytes([next[0], next[1]]);
            if !(0xdc00..=0xdfff).contains(&low) {
                return Err(BindingError::InvalidUtf16 {
                    offset: absolute_data_start + index,
                });
            }
            index = next_index + 2;
            let high = u32::from(unit) - 0xd800;
            let low = u32::from(low) - 0xdc00;
            char::from_u32(0x1_0000 + (high << 10) + low).ok_or(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index - 4,
            })?
        } else if (0xdc00..=0xdfff).contains(&unit) {
            return Err(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index,
            });
        } else {
            index = index.checked_add(2).ok_or(BindingError::LengthOverflow)?;
            char::from_u32(u32::from(unit)).ok_or(BindingError::InvalidUtf16 {
                offset: absolute_data_start + index - 2,
            })?
        };
        bytes = bytes
            .checked_add(value.len_utf8())
            .ok_or(BindingError::LengthOverflow)?;
    }
    let _ = field;
    Ok(bytes)
}

fn parse_connection_header(record: RawRecord<'_>) -> Result<(u32, u32)> {
    let payload = record.payload_bytes;
    let source_type = read_u32(payload, 10)?;
    let connection_id = read_u32(payload, 18)?;
    Ok((source_type, connection_id))
}

fn parse_frt_begin(record: RawRecord<'_>) -> Result<FrtProductVersion> {
    if record.payload_bytes.len() != 4 {
        return Err(BindingError::InvalidWrapperPayload {
            kind: FRT_BEGIN,
            offset: record.span.payload.start(),
        });
    }
    let version = read_u16(record.payload_bytes, 0)?;
    let product_bits = read_u16(record.payload_bytes, 2)?;
    if product_bits & 0x8000 != 0 {
        return Err(BindingError::InvalidWrapperPayload {
            kind: FRT_BEGIN,
            offset: record.span.payload.start() + 2,
        });
    }
    Ok(FrtProductVersion {
        version,
        product: product_bits & 0x7fff,
    })
}

fn parse_ac_begin<B: WorkBudget + ?Sized>(
    record: RawRecord<'_>,
    limits: &BindingLimits,
    work: &mut usize,
    budget: &mut B,
) -> Result<Vec<AcProductVersion>> {
    if record.payload_bytes.len() < 2 {
        return Err(BindingError::InvalidWrapperPayload {
            kind: AC_BEGIN,
            offset: record.span.payload.start(),
        });
    }
    let count = usize::from(read_u16(record.payload_bytes, 0)?);
    if count == 0 {
        return Err(BindingError::InvalidWrapperPayload {
            kind: AC_BEGIN,
            offset: record.span.payload.start(),
        });
    }
    let expected = 2usize
        .checked_add(count.checked_mul(4).ok_or(BindingError::LengthOverflow)?)
        .ok_or(BindingError::LengthOverflow)?;
    if expected != record.payload_bytes.len() {
        return Err(BindingError::InvalidWrapperPayload {
            kind: AC_BEGIN,
            offset: record.span.payload.start(),
        });
    }
    if count > limits.max_records {
        return Err(BindingError::RecordLimit {
            actual: count,
            limit: limits.max_records,
        });
    }
    let mut versions = Vec::new();
    let reservation = count
        .checked_mul(size_of::<AcProductVersion>())
        .ok_or(BindingError::LengthOverflow)?;
    charge_scan_work(work, reservation, limits, budget)?;
    try_reserve(&mut versions, count, "alternate-content product versions")?;
    for index in 0..count {
        let offset = 2 + index * 4;
        let file_version = read_u16(record.payload_bytes, offset)?;
        let bits = read_u16(record.payload_bytes, offset + 2)?;
        versions.push(AcProductVersion {
            file_version,
            file_product: bits & 0x7fff,
            extension: bits & 0x8000 != 0,
        });
    }
    Ok(versions)
}

fn frt_wrapper(
    begin: RecordSpan,
    end: RecordSpan,
    profile: Option<FrtProductVersion>,
) -> Result<FrtWrapperSpan> {
    let profile = profile.ok_or(BindingError::UnsupportedWrapper {
        offset: begin.range.start(),
        detail: "malformed FRTProductVersion",
    })?;
    Ok(FrtWrapperSpan {
        begin,
        end,
        range: Span::new(begin.range.start(), end.range.end())?,
        profile,
    })
}

fn next_record<'a>(
    source: &'a [u8],
    cursor: &mut usize,
    limits: &BindingLimits,
) -> Result<RawRecord<'a>> {
    let start = *cursor;
    let first = *source.get(start).ok_or(BindingError::UnexpectedEnd {
        context: "BIFF12 record kind",
    })?;
    let (kind, kind_len) = if first & 0x80 == 0 {
        (u16::from(first), 1usize)
    } else {
        let second = *source.get(start + 1).ok_or(BindingError::Truncated {
            offset: start + 1,
            needed: 1,
            available: source.len().saturating_sub(start + 1),
            context: "BIFF12 record kind",
        })?;
        if second & 0x80 != 0 {
            return Err(BindingError::InvalidKind { offset: start });
        }
        let value = u16::from(first & 0x7f) | (u16::from(second) << 7);
        if value < 0x80 {
            return Err(BindingError::InvalidKind { offset: start });
        }
        (value, 2usize)
    };
    let len_start = start
        .checked_add(kind_len)
        .ok_or(BindingError::LengthOverflow)?;
    let mut payload_len = 0u32;
    let mut length_len = 0usize;
    for index in 0..4usize {
        let position = len_start
            .checked_add(index)
            .ok_or(BindingError::LengthOverflow)?;
        let byte = *source.get(position).ok_or(BindingError::Truncated {
            offset: position,
            needed: 1,
            available: source.len().saturating_sub(position),
            context: "BIFF12 record payload length",
        })?;
        if index == 3 && byte & 0x80 != 0 {
            return Err(BindingError::InvalidLengthVarint { offset: len_start });
        }
        payload_len |= u32::from(byte & 0x7f) << (index * 7);
        length_len = index + 1;
        if byte & 0x80 == 0 {
            break;
        }
    }
    let payload_len =
        usize::try_from(payload_len).map_err(|_error| BindingError::LengthOverflow)?;
    if payload_len > limits.max_record_payload_bytes {
        return Err(BindingError::RecordPayloadLimit {
            actual: payload_len,
            limit: limits.max_record_payload_bytes,
        });
    }
    let payload_start = len_start
        .checked_add(length_len)
        .ok_or(BindingError::LengthOverflow)?;
    let payload_end = payload_start
        .checked_add(payload_len)
        .ok_or(BindingError::LengthOverflow)?;
    let payload_bytes = source
        .get(payload_start..payload_end)
        .ok_or(BindingError::Truncated {
            offset: payload_start,
            needed: payload_len,
            available: source.len().saturating_sub(payload_start),
            context: "BIFF12 record payload",
        })?;
    *cursor = payload_end;
    Ok(RawRecord {
        span: RecordSpan {
            range: Span::new(start, payload_end)?,
            payload: Span::new(payload_start, payload_end)?,
            kind,
            kind_len: u8::try_from(kind_len).map_err(|_error| BindingError::LengthOverflow)?,
            length_len: u8::try_from(length_len).map_err(|_error| BindingError::LengthOverflow)?,
        },
        payload_bytes,
    })
}

fn read_u16(payload: &[u8], offset: usize) -> Result<u16> {
    let bytes = payload
        .get(offset..offset.checked_add(2).ok_or(BindingError::LengthOverflow)?)
        .ok_or(BindingError::Truncated {
            offset,
            needed: 2,
            available: payload.len().saturating_sub(offset),
            context: "fixed u16",
        })?;
    Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
}

fn read_u32(payload: &[u8], offset: usize) -> Result<u32> {
    let bytes = payload
        .get(offset..offset.checked_add(4).ok_or(BindingError::LengthOverflow)?)
        .ok_or(BindingError::Truncated {
            offset,
            needed: 4,
            available: payload.len().saturating_sub(offset),
            context: "fixed u32",
        })?;
    Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
}

fn require_empty_payload(record: RawRecord<'_>, context: &'static str) -> Result<()> {
    if !record.payload_bytes.is_empty() {
        return Err(BindingError::TrailingPayload {
            offset: record.span.payload.start(),
            remaining: record.payload_bytes.len(),
            context,
        });
    }
    Ok(())
}

fn push_binding<B: WorkBudget + ?Sized>(
    bindings: &mut Vec<ExtConn14Binding>,
    binding: ExtConn14Binding,
    limits: &BindingLimits,
    work: &mut usize,
    budget: &mut B,
) -> Result<()> {
    if bindings.len() >= limits.max_bindings {
        return Err(BindingError::BindingLimit {
            actual: bindings.len() + 1,
            limit: limits.max_bindings,
        });
    }
    if bindings.len() == bindings.capacity() {
        charge_scan_work(work, size_of::<ExtConn14Binding>(), limits, budget)?;
        bindings
            .try_reserve_exact(1)
            .map_err(|_error| BindingError::Allocation {
                resource: "ExtConn14 bindings",
            })?;
    }
    bindings.push(binding);
    Ok(())
}

fn push_opaque<B: WorkBudget + ?Sized>(
    blocks: &mut Vec<OpaqueBlock>,
    block: OpaqueBlock,
    limits: &BindingLimits,
    work: &mut usize,
    budget: &mut B,
) -> Result<()> {
    if blocks.len() >= limits.max_opaque_blocks {
        return Err(BindingError::OpaqueLimit {
            actual: blocks.len() + 1,
            limit: limits.max_opaque_blocks,
        });
    }
    if blocks.len() == blocks.capacity() {
        charge_scan_work(work, size_of::<OpaqueBlock>(), limits, budget)?;
        blocks
            .try_reserve_exact(1)
            .map_err(|_error| BindingError::Allocation {
                resource: "opaque wrapper provenance",
            })?;
    }
    blocks.push(block);
    Ok(())
}

fn try_reserve<T>(vector: &mut Vec<T>, additional: usize, resource: &'static str) -> Result<()> {
    vector
        .try_reserve_exact(additional)
        .map_err(|_error| BindingError::Allocation { resource })
}

fn validate_binding_spans(bindings: &[ExtConn14Binding], source_len: usize) -> Result<()> {
    let mut previous: Option<Span> = None;
    for binding in bindings {
        let record = binding.record.range;
        if record.end() > source_len
            || binding.client_cube_urn_span.field.end() > source_len
            || !record.contains(binding.client_cube_urn_span.field)
            || !record.contains(binding.culture_span.field)
            || binding
                .client_cube_urn_span
                .field
                .overlaps(binding.culture_span.field)
        {
            return Err(BindingError::DuplicateBindingSpan);
        }
        if let Some(previous) = previous {
            if previous.overlaps(record) || previous.end() > record.start() {
                return Err(BindingError::DuplicateBindingSpan);
            }
        }
        previous = Some(record);
        let wrapper = binding.wrapper.range;
        if !wrapper.contains(record) {
            return Err(BindingError::DuplicateBindingSpan);
        }
    }
    Ok(())
}

fn validate_opaque_spans(blocks: &[OpaqueBlock], source_len: usize) -> Result<()> {
    for block in blocks {
        if block.range.end() > source_len {
            return Err(BindingError::SourceSpanOutOfBounds {
                start: block.range.start(),
                end: block.range.end(),
                source_len,
            });
        }
    }
    Ok(())
}

fn opaque_blocks_have_candidates(blocks: &[OpaqueBlock]) -> bool {
    blocks
        .iter()
        .any(|block| block.contains_reference_candidate)
}

fn validate_replacement_order(replacements: &[PlannedReplacement<'_>]) -> Result<()> {
    let mut previous: Option<Span> = None;
    for replacement in replacements {
        if replacement.record.payload.start() > replacement.string.field.start()
            || !replacement
                .record
                .payload
                .contains(replacement.string.field)
        {
            return Err(BindingError::DuplicateBindingSpan);
        }
        if let Some(previous) = previous {
            if previous.overlaps(replacement.record.range)
                || previous.end() > replacement.record.range.start()
            {
                return Err(BindingError::DuplicateBindingSpan);
            }
        }
        previous = Some(replacement.record.range);
    }
    Ok(())
}

fn wire_wide_string_len_from_units(units: usize) -> Result<usize> {
    units
        .checked_mul(2)
        .and_then(|value| value.checked_add(4))
        .ok_or(BindingError::LengthOverflow)
}

fn encoded_varint_len(mut value: usize) -> usize {
    let mut len = 1usize;
    while value >= 0x80 {
        value >>= 7;
        len += 1;
    }
    len
}

fn encode_varint(value: usize, output: &mut Vec<u8>) -> Result<()> {
    if value > MAX_WIRE_PAYLOAD {
        return Err(BindingError::LengthOverflow);
    }
    let mut value = value;
    loop {
        let mut byte = u8::try_from(value & 0x7f).map_err(|_error| BindingError::LengthOverflow)?;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return Ok(());
        }
    }
}

/// A typed bounded failure from the scanner or rewrite planner.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub(crate) enum BindingError {
    /// A span constructor received reversed bounds.
    InvalidSpan { start: usize, end: usize },
    /// Caller supplied a ceiling outside this profile's representable domain.
    InvalidLimit {
        /// Limit field name.
        name: &'static str,
        /// Supplied value.
        value: usize,
        /// Hard maximum.
        maximum: usize,
    },
    /// Source part exceeded its caller ceiling.
    SourceLimit { actual: usize, limit: usize },
    /// Framed record count exceeded its caller ceiling.
    RecordLimit { actual: usize, limit: usize },
    /// Effective reference count exceeded its caller ceiling.
    BindingLimit { actual: usize, limit: usize },
    /// Connection context count overflowed its bounded ordinal.
    ConnectionLimit { actual: usize, limit: usize },
    /// Opaque provenance count exceeded its caller ceiling.
    OpaqueLimit { actual: usize, limit: usize },
    /// ExtConn14 calculated-member record count exceeded its caller ceiling.
    CalculatedMemberLimit { actual: usize, limit: usize },
    /// Nested wrapper count exceeded its caller ceiling.
    WrapperDepthLimit { limit: usize },
    /// UTF-16 string exceeded its hard or caller ceiling.
    StringLimit {
        /// Field name.
        field: &'static str,
        /// Decoded code units.
        units: usize,
        /// Effective limit.
        limit: usize,
        /// Length-prefix source offset.
        offset: usize,
    },
    /// A record payload exceeded its caller ceiling.
    RecordPayloadLimit { actual: usize, limit: usize },
    /// A rewrite output exceeded its caller ceiling.
    OutputLimit { actual: usize, limit: usize },
    /// Planned source/output work exceeded its caller ceiling.
    WorkLimit { actual: usize, limit: usize },
    /// A cancellation or caller budget rejected the operation.
    BudgetRejected,
    /// The BIFF12 kind varint was malformed.
    InvalidKind { offset: usize },
    /// The BIFF12 payload-length varint was malformed or non-canonical.
    InvalidLengthVarint { offset: usize },
    /// A source field or record ended before the pinned grammar completed.
    Truncated {
        /// Source offset.
        offset: usize,
        /// Required bytes.
        needed: usize,
        /// Available bytes.
        available: usize,
        /// Field/record context.
        context: &'static str,
    },
    /// UTF-16LE contained an unpaired surrogate.
    InvalidUtf16 { offset: usize },
    /// A wrapper payload did not match its normative fixed grammar.
    InvalidWrapperPayload { kind: u16, offset: usize },
    /// A fixed-width payload had bytes after its grammar.
    TrailingPayload {
        /// First trailing byte.
        offset: usize,
        /// Trailing count.
        remaining: usize,
        /// Record context.
        context: &'static str,
    },
    /// A record kind was not legal at its current envelope position.
    UnexpectedRecord {
        /// Source offset.
        offset: usize,
        /// Expected kind.
        expected: u16,
        /// Actual kind.
        found: u16,
    },
    /// A stream ended before a required root or collection end.
    UnexpectedEnd { context: &'static str },
    /// A known collection end had no matching begin.
    UnbalancedWrapper { kind: u16, offset: usize },
    /// An ExtConn14 candidate was outside the admitted wrapper/profile.
    UnsupportedWrapper { offset: usize, detail: &'static str },
    /// The DBTOLEDB/CMDCUBE/PivotCache proof was absent or contradictory.
    Context { offset: usize, detail: &'static str },
    /// A possible reference was hidden in opaque AC/FRT/ExtConn15 bytes.
    HiddenReference,
    /// The same physical source range was decoded more than once.
    DuplicateBindingSpan,
    /// The source allocation did not match the scan that produced the spans.
    SourceMismatch,
    /// The edit list named one source UID more than once with the same target.
    DuplicateRewrite,
    /// The edit list named one source UID with conflicting targets.
    ConflictingRewrite,
    /// Replacement UID exceeded the configured UTF-16 cap.
    UidLimit { units: usize, limit: usize },
    /// An edit list exceeded its caller ceiling.
    RewriteCountLimit { actual: usize, limit: usize },
    /// A checked arithmetic operation overflowed.
    LengthOverflow,
    /// Fallible vector/string allocation was rejected.
    Allocation { resource: &'static str },
    /// Internal state invariant was not preserved.
    StateCorrupt,
    /// Non-record bytes followed the root end.
    TrailingRecords { offset: usize },
    /// A source span did not lie within the source allocation.
    SourceSpanOutOfBounds {
        /// Span start.
        start: usize,
        /// Span end.
        end: usize,
        /// Source allocation length.
        source_len: usize,
    },
}

impl fmt::Display for BindingError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidSpan { start, end } => {
                write!(formatter, "invalid source span {start}..{end}")
            },
            Self::InvalidLimit {
                name,
                value,
                maximum,
            } => write!(
                formatter,
                "invalid {name} limit {value}; maximum is {maximum}"
            ),
            Self::SourceLimit { actual, limit } => {
                write!(formatter, "source bytes {actual} exceed limit {limit}")
            },
            Self::RecordLimit { actual, limit } => {
                write!(formatter, "record count {actual} exceeds limit {limit}")
            },
            Self::BindingLimit { actual, limit } => {
                write!(formatter, "binding count {actual} exceeds limit {limit}")
            },
            Self::ConnectionLimit { actual, limit } => {
                write!(formatter, "connection count {actual} exceeds limit {limit}")
            },
            Self::OpaqueLimit { actual, limit } => {
                write!(
                    formatter,
                    "opaque block count {actual} exceeds limit {limit}"
                )
            },
            Self::CalculatedMemberLimit { actual, limit } => write!(
                formatter,
                "calculated-member record count {actual} exceeds limit {limit}"
            ),
            Self::WrapperDepthLimit { limit } => {
                write!(formatter, "wrapper depth exceeds limit {limit}")
            },
            Self::StringLimit {
                field,
                units,
                limit,
                offset,
            } => write!(
                formatter,
                "{field} UTF-16 units {units} exceed limit {limit} at byte {offset}"
            ),
            Self::RecordPayloadLimit { actual, limit } => {
                write!(formatter, "record payload {actual} exceeds limit {limit}")
            },
            Self::OutputLimit { actual, limit } => {
                write!(formatter, "rewrite output {actual} exceeds limit {limit}")
            },
            Self::WorkLimit { actual, limit } => {
                write!(formatter, "planned work {actual} exceeds limit {limit}")
            },
            Self::BudgetRejected => formatter.write_str("caller work budget rejected operation"),
            Self::InvalidKind { offset } => {
                write!(formatter, "invalid BIFF12 kind encoding at byte {offset}")
            },
            Self::InvalidLengthVarint { offset } => {
                write!(
                    formatter,
                    "invalid BIFF12 payload-length varint at byte {offset}"
                )
            },
            Self::Truncated {
                offset,
                needed,
                available,
                context,
            } => write!(
                formatter,
                "truncated {context} at byte {offset}: needed {needed}, found {available}"
            ),
            Self::InvalidUtf16 { offset } => write!(formatter, "invalid UTF-16LE at byte {offset}"),
            Self::InvalidWrapperPayload { kind, offset } => {
                write!(
                    formatter,
                    "invalid wrapper payload for kind {kind} at byte {offset}"
                )
            },
            Self::TrailingPayload {
                offset,
                remaining,
                context,
            } => write!(
                formatter,
                "{context} has {remaining} trailing bytes at byte {offset}"
            ),
            Self::UnexpectedRecord {
                offset,
                expected,
                found,
            } => write!(
                formatter,
                "unexpected record kind {found} at byte {offset}; expected {expected}"
            ),
            Self::UnexpectedEnd { context } => write!(formatter, "unexpected end of {context}"),
            Self::UnbalancedWrapper { kind, offset } => {
                write!(formatter, "unbalanced wrapper kind {kind} at byte {offset}")
            },
            Self::UnsupportedWrapper { offset, detail } => {
                write!(formatter, "unsupported wrapper at byte {offset}: {detail}")
            },
            Self::Context { offset, detail } => {
                write!(
                    formatter,
                    "invalid ExtConn14 context at byte {offset}: {detail}"
                )
            },
            Self::HiddenReference => {
                formatter.write_str("opaque source may hide a Custom Data reference")
            },
            Self::DuplicateBindingSpan => {
                formatter.write_str("duplicate or overlapping binding span")
            },
            Self::SourceMismatch => formatter.write_str("source does not match the captured scan"),
            Self::DuplicateRewrite => formatter.write_str("duplicate rewrite for one UID"),
            Self::ConflictingRewrite => formatter.write_str("conflicting rewrites for one UID"),
            Self::UidLimit { units, limit } => {
                write!(
                    formatter,
                    "replacement UID units {units} exceed limit {limit}"
                )
            },
            Self::RewriteCountLimit { actual, limit } => {
                write!(formatter, "rewrite count {actual} exceeds limit {limit}")
            },
            Self::LengthOverflow => formatter.write_str("checked length arithmetic overflow"),
            Self::Allocation { resource } => {
                write!(formatter, "allocation rejected for {resource}")
            },
            Self::StateCorrupt => formatter.write_str("binding scanner state invariant failed"),
            Self::TrailingRecords { offset } => {
                write!(
                    formatter,
                    "records follow BrtEndExtConnections at byte {offset}"
                )
            },
            Self::SourceSpanOutOfBounds {
                start,
                end,
                source_len,
            } => write!(
                formatter,
                "source span {start}..{end} exceeds source length {source_len}"
            ),
        }
    }
}

impl std::error::Error for BindingError {}

/// Result alias for scanner and rewrite operations.
pub(crate) type Result<T> = std::result::Result<T, BindingError>;

#[cfg(test)]
mod tests {
    use super::*;

    struct AllowBudget;

    impl WorkBudget for AllowBudget {}

    #[derive(Default)]
    struct RejectBudget {
        calls: usize,
        reject_after: usize,
    }

    impl WorkBudget for RejectBudget {
        fn charge(&mut self, _units: usize) -> Result<()> {
            self.calls += 1;
            if self.calls > self.reject_after {
                Err(BindingError::BudgetRejected)
            } else {
                Ok(())
            }
        }
    }

    fn test_limits() -> BindingLimits {
        BindingLimits {
            max_source_bytes: 1 << 20,
            max_record_payload_bytes: 1 << 20,
            max_records: 1 << 12,
            max_bindings: 1 << 10,
            max_opaque_blocks: 1 << 10,
            max_wrapper_depth: 64,
            max_calculated_member_records: 1 << 10,
            max_culture_units: HARD_CULTURE_UNITS,
            max_client_cube_urn_units: HARD_CLIENT_CUBE_URN_UNITS,
            max_output_bytes: 1 << 20,
            max_work_bytes: 1 << 24,
            initial_binding_capacity: 16,
            initial_opaque_capacity: 8,
        }
    }

    fn varint(mut value: usize, output: &mut Vec<u8>) {
        loop {
            let mut byte = match u8::try_from(value & 0x7f) {
                Ok(byte) => byte,
                Err(_) => unreachable!("varint byte is masked to seven bits"),
            };
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

    fn record(kind: u16, payload: &[u8], output: &mut Vec<u8>) {
        varint(kind as usize, output);
        varint(payload.len(), output);
        output.extend_from_slice(payload);
    }

    fn wide(value: &str, output: &mut Vec<u8>) {
        let units: Vec<u16> = value.encode_utf16().collect();
        let unit_count = match u32::try_from(units.len()) {
            Ok(unit_count) => unit_count,
            Err(_) => unreachable!("synthetic test string length fits in u32"),
        };
        output.extend_from_slice(&unit_count.to_le_bytes());
        for unit in units {
            output.extend_from_slice(&unit.to_le_bytes());
        }
    }

    fn ext_connection(id: u32) -> Vec<u8> {
        let mut payload = vec![7, 5, 2, 0];
        payload.extend_from_slice(&30_u16.to_le_bytes());
        payload.extend_from_slice(&0_u16.to_le_bytes());
        payload.extend_from_slice(&(1_u16 << 3).to_le_bytes());
        payload.extend_from_slice(&5_u32.to_le_bytes());
        payload.extend_from_slice(&3_u32.to_le_bytes());
        payload.extend_from_slice(&id.to_le_bytes());
        payload.push(0);
        wide("Synthetic", &mut payload);
        payload
    }

    fn ext14(uid: &str) -> Vec<u8> {
        let mut payload = vec![1, 2, 3, 4];
        wide("en-US", &mut payload);
        wide(uid, &mut payload);
        payload
    }

    fn source() -> Vec<u8> {
        let mut output = Vec::new();
        record(BEGIN_EXT_CONNECTIONS, &[], &mut output);
        record(BEGIN_EXT_CONNECTION, &ext_connection(42), &mut output);
        for uid in ["uid-A", "uid-A"] {
            record(FRT_BEGIN, &[1, 0, 1, 0], &mut output);
            record(BEGIN_EXT_CONN14, &ext14(uid), &mut output);
            record(END_EXT_CONN14, &[], &mut output);
            record(FRT_END, &[], &mut output);
        }
        record(END_EXT_CONNECTION, &[], &mut output);
        record(END_EXT_CONNECTIONS, &[], &mut output);
        output
    }

    fn source_without_connection() -> Vec<u8> {
        let mut output = Vec::new();
        record(BEGIN_EXT_CONNECTIONS, &[], &mut output);
        record(FRT_BEGIN, &[1, 0, 1, 0], &mut output);
        record(BEGIN_EXT_CONN14, &ext14("uid-A"), &mut output);
        record(END_EXT_CONN14, &[], &mut output);
        record(FRT_END, &[], &mut output);
        record(END_EXT_CONNECTIONS, &[], &mut output);
        output
    }

    fn source_with_nested_frt() -> Vec<u8> {
        let mut output = Vec::new();
        record(BEGIN_EXT_CONNECTIONS, &[], &mut output);
        record(BEGIN_EXT_CONNECTION, &ext_connection(42), &mut output);
        record(FRT_BEGIN, &[1, 0, 1, 0], &mut output);
        record(BEGIN_EXT_CONN14, &ext14("uid-A"), &mut output);
        record(FRT_BEGIN, &[1, 0, 1, 0], &mut output);
        record(1, &[], &mut output);
        record(FRT_END, &[], &mut output);
        record(END_EXT_CONN14, &[], &mut output);
        record(FRT_END, &[], &mut output);
        record(END_EXT_CONNECTION, &[], &mut output);
        record(END_EXT_CONNECTIONS, &[], &mut output);
        output
    }

    #[test]
    fn nonzero_reserved_blank_and_many_to_one_batch_rewrite() {
        let source = source();
        let mut budget = AllowBudget;
        let scan = scan_connections_with_budget(&source, test_limits(), &[], &mut budget).unwrap();
        assert!(scan.record_count() >= scan.bindings().len());
        assert_eq!(scan.bindings().len(), 2);
        assert_eq!(scan.bindings()[0].client_cube_urn(), "uid-A");
        assert_eq!(scan.bindings()[1].client_cube_urn(), "uid-A");
        let mut output = Vec::new();
        let mut budget = AllowBudget;
        let outcome = rewrite_client_cube_urns_with_budget(
            &source,
            &scan,
            &[UidRewrite {
                from: "uid-A",
                to: "uid-B",
            }],
            &mut output,
            test_limits(),
            &mut budget,
        )
        .unwrap();
        assert!(outcome.changed);
        assert_eq!(outcome.replacements, 2);
        let mut budget = AllowBudget;
        let rewritten =
            scan_connections_with_budget(&output, test_limits(), &[], &mut budget).unwrap();
        assert!(
            rewritten
                .bindings()
                .iter()
                .all(|binding| binding.client_cube_urn() == "uid-B")
        );
    }

    #[test]
    fn duplicate_edits_are_bounded_and_typed() {
        let source = source();
        let mut budget = AllowBudget;
        let scan = scan_connections_with_budget(&source, test_limits(), &[], &mut budget).unwrap();
        let mut output = vec![0xa5, 0x5a];
        let before = output.clone();
        let mut budget = AllowBudget;
        let error = rewrite_client_cube_urns_with_budget(
            &source,
            &scan,
            &[
                UidRewrite {
                    from: "uid-A",
                    to: "uid-B",
                },
                UidRewrite {
                    from: "uid-A",
                    to: "uid-B",
                },
            ],
            &mut output,
            test_limits(),
            &mut budget,
        )
        .unwrap_err();
        assert_eq!(error, BindingError::DuplicateRewrite);
        assert_eq!(output, before);
    }

    #[test]
    fn budget_refusal_happens_before_output_mutation() {
        let source = source();
        let mut budget = AllowBudget;
        let scan = scan_connections_with_budget(&source, test_limits(), &[], &mut budget).unwrap();
        let mut output = vec![0x7f, 0x01, 0x7f];
        let before = output.clone();
        let mut budget = RejectBudget {
            reject_after: 0,
            ..RejectBudget::default()
        };
        let error = rewrite_client_cube_urns_with_budget(
            &source,
            &scan,
            &[UidRewrite {
                from: "uid-A",
                to: "uid-B",
            }],
            &mut output,
            test_limits(),
            &mut budget,
        )
        .unwrap_err();
        assert_eq!(error, BindingError::BudgetRejected);
        assert_eq!(output, before);
    }

    #[test]
    fn unsupported_context_and_nested_future_wrapper_are_full_opaque_ranges() {
        let no_connection_source = source_without_connection();
        let mut budget = AllowBudget;
        let no_connection =
            scan_connections_with_budget(&no_connection_source, test_limits(), &[], &mut budget)
                .unwrap();
        assert!(no_connection.has_blocking_opaque());
        assert!(no_connection.bindings().is_empty());
        assert_eq!(no_connection.opaque_blocks[0].range.start(), 9);
        assert!(no_connection.opaque_blocks[0].range.end() > 9);

        let nested_source = source_with_nested_frt();
        let mut budget = AllowBudget;
        let nested =
            scan_connections_with_budget(&nested_source, test_limits(), &[], &mut budget).unwrap();
        assert!(nested.has_blocking_opaque());
        assert!(nested.bindings().is_empty());
        assert!(
            nested
                .opaque_blocks
                .iter()
                .any(|block| block.range.end() > block.range.start())
        );
    }

    #[test]
    fn rewrite_rechecks_untouched_record_payload_limits() {
        let mut source = source();
        let extra = vec![0x5a; 100];
        let end = source.len() - 3;
        let mut record_bytes = Vec::new();
        record(7, &extra, &mut record_bytes);
        source.splice(end..end, record_bytes);
        let mut budget = AllowBudget;
        let scan = scan_connections_with_budget(&source, test_limits(), &[], &mut budget).unwrap();
        let mut output = vec![0x44, 0x55];
        let before = output.clone();
        let mut limits = test_limits();
        limits.max_record_payload_bytes = 18;
        let mut budget = AllowBudget;
        let error = rewrite_client_cube_urns_with_budget(
            &source,
            &scan,
            &[UidRewrite {
                from: "uid-A",
                to: "uid-B",
            }],
            &mut output,
            limits,
            &mut budget,
        )
        .unwrap_err();
        assert!(matches!(error, BindingError::RecordPayloadLimit { .. }));
        assert_eq!(output, before);
    }

    #[test]
    fn begin_connection_inside_frt_is_retained_as_opaque() {
        let mut source = Vec::new();
        record(BEGIN_EXT_CONNECTIONS, &[], &mut source);
        record(FRT_BEGIN, &[1, 0, 1, 0], &mut source);
        record(BEGIN_EXT_CONNECTION, &ext_connection(42), &mut source);
        record(END_EXT_CONNECTION, &[], &mut source);
        record(FRT_END, &[], &mut source);
        record(END_EXT_CONNECTIONS, &[], &mut source);
        let mut budget = AllowBudget;
        let scan = scan_connections_with_budget(&source, test_limits(), &[], &mut budget).unwrap();
        assert!(scan.has_blocking_opaque());
        assert!(scan.bindings().is_empty());
        assert_eq!(scan.opaque_blocks.len(), 1);
        assert_eq!(scan.opaque_blocks[0].range.start(), 3);
        assert_eq!(scan.opaque_blocks[0].range.end(), source.len() - 3);
    }
}
