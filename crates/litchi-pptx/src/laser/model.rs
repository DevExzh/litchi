//! Semantic laser-trace values and bounded resource state.

use crate::time::Offset;
use litchi_drawingml::coordinate::Coordinate;

pub(super) const PRESENTATIONML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(super) const STRICT_PRESENTATIONML_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/presentationml/main";

/// The `PresentationML` namespace dialect used by a laser-trace writer.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Conformance {
    /// The ISO Transitional `PresentationML` namespace.
    #[default]
    Transitional,
    /// The ISO Strict `PresentationML` namespace.
    Strict,
}

impl Conformance {
    /// Return the `PresentationML` namespace URI used by this profile.
    #[inline]
    #[must_use]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => PRESENTATIONML_NAMESPACE,
            Self::Strict => STRICT_PRESENTATIONML_NAMESPACE,
        }
    }

    /// Select the profile for a detected `PresentationML` namespace.
    #[inline]
    #[must_use]
    pub fn from_namespace(namespace: &str) -> Self {
        if namespace == STRICT_PRESENTATIONML_NAMESPACE {
            Self::Strict
        } else {
            Self::Transitional
        }
    }
}

/// A persisted laser-pointer point from a `PowerPoint` slide show.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TracePoint {
    pub(super) time: Offset,
    pub(super) x: Coordinate,
    pub(super) y: Coordinate,
}

impl TracePoint {
    /// Create a trace point from exact, checked time and coordinate values.
    #[must_use]
    pub fn new(time: Offset, x: Coordinate, y: Coordinate) -> Self {
        Self { time, x, y }
    }

    /// Return the exact normalized time offset relative to the slide timeline.
    #[inline]
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the checked horizontal `DrawingML` coordinate.
    #[inline]
    pub fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Return the checked vertical `DrawingML` coordinate.
    #[inline]
    pub fn y(&self) -> &Coordinate {
        &self.y
    }
}

/// A bounded, inert laser-pointer trace recorded for a presentation slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Trace {
    pub(super) slide_index: usize,
    pub(super) trace_index: usize,
    pub(super) points: Vec<TracePoint>,
}

impl Trace {
    /// Return the zero-based index of the slide that owns this trace.
    #[inline]
    #[must_use]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this trace on its slide.
    #[inline]
    #[must_use]
    pub fn trace_index(&self) -> usize {
        self.trace_index
    }

    /// Return the stored trace points in source order.
    #[inline]
    #[must_use]
    pub fn points(&self) -> &[TracePoint] {
        &self.points
    }

    /// Return the number of stored trace points.
    #[inline]
    #[must_use]
    pub fn point_count(&self) -> usize {
        self.points.len()
    }
}

/// Caller-selected resource policy for laser-trace XML operations.
///
/// This policy is immutable.  Aggregate accounting for a package or a slide
/// walk lives in [`Usage`], so tightening a policy cannot accidentally reset
/// or reuse a mutable counter.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    max_slide_xml_bytes: usize,
    max_total_slide_xml_bytes: usize,
    max_trace_count: usize,
    max_point_count: usize,
    max_xml_nodes: usize,
    max_xml_depth: usize,
    max_output_bytes: usize,
}

impl Limits {
    /// Conservative bounded policy for ordinary PresentationML slides.
    pub const DEFAULT: Self = Self {
        max_slide_xml_bytes: 32 * 1024 * 1024,
        max_total_slide_xml_bytes: 256 * 1024 * 1024,
        max_trace_count: 4_096,
        max_point_count: 65_536,
        max_xml_nodes: 250_000,
        max_xml_depth: 128,
        max_output_bytes: 32 * 1024 * 1024,
    };

    /// Construct a finite policy, bounded by the crate's safety ceilings.
    #[must_use]
    pub const fn new(
        max_slide_xml_bytes: usize,
        max_total_slide_xml_bytes: usize,
        max_trace_count: usize,
        max_point_count: usize,
        max_xml_nodes: usize,
        max_xml_depth: usize,
        max_output_bytes: usize,
    ) -> Option<Self> {
        let default = Self::DEFAULT;
        if max_slide_xml_bytes == 0
            || max_slide_xml_bytes > default.max_slide_xml_bytes
            || max_total_slide_xml_bytes == 0
            || max_total_slide_xml_bytes > default.max_total_slide_xml_bytes
            || max_trace_count == 0
            || max_trace_count > default.max_trace_count
            || max_point_count == 0
            || max_point_count > default.max_point_count
            || max_xml_nodes == 0
            || max_xml_nodes > default.max_xml_nodes
            || max_xml_depth == 0
            || max_xml_depth > default.max_xml_depth
            || max_output_bytes == 0
            || max_output_bytes > default.max_output_bytes
        {
            return None;
        }
        Some(Self {
            max_slide_xml_bytes,
            max_total_slide_xml_bytes,
            max_trace_count,
            max_point_count,
            max_xml_nodes,
            max_xml_depth,
            max_output_bytes,
        })
    }

    /// Maximum bytes accepted from one slide XML part.
    #[must_use]
    pub const fn max_slide_xml_bytes(self) -> usize {
        self.max_slide_xml_bytes
    }

    /// Maximum aggregate bytes accepted across slide XML parts.
    #[must_use]
    pub const fn max_total_slide_xml_bytes(self) -> usize {
        self.max_total_slide_xml_bytes
    }

    /// Maximum traces retained across the operation's usage ledger.
    #[must_use]
    pub const fn max_trace_count(self) -> usize {
        self.max_trace_count
    }

    /// Maximum points retained across the operation's usage ledger.
    #[must_use]
    pub const fn max_point_count(self) -> usize {
        self.max_point_count
    }

    /// Maximum XML nodes scanned in one slide.
    #[must_use]
    pub const fn max_xml_nodes(self) -> usize {
        self.max_xml_nodes
    }

    /// Maximum XML nesting depth scanned in one slide.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.max_xml_depth
    }

    /// Maximum bytes emitted by one trace fragment.
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    /// Return this policy with a separate bounded output budget.
    #[must_use]
    pub const fn with_output_bytes(mut self, max_output_bytes: usize) -> Option<Self> {
        if max_output_bytes == 0 || max_output_bytes > Self::DEFAULT.max_output_bytes {
            return None;
        }
        self.max_output_bytes = max_output_bytes;
        Some(self)
    }

    /// Intersect the codec policy with the retained OPC owner policy.
    ///
    /// Package edits must use one such intersection for every preflight,
    /// scan, write, insertion, and read-back step.  The codec ceilings remain
    /// hard safety ceilings when the owner was admitted under a broader
    /// source profile.
    pub(crate) fn from_source_limits(source: litchi_opc::SourceLimits) -> Self {
        let default = Self::DEFAULT;
        let part = usize::try_from(source.max_part_bytes()).unwrap_or(usize::MAX);
        let entry = usize::try_from(source.max_entry_bytes()).unwrap_or(usize::MAX);
        let total_part = usize::try_from(source.max_total_part_bytes()).unwrap_or(usize::MAX);
        let aggregate = usize::try_from(source.max_aggregate_bytes()).unwrap_or(usize::MAX);
        Self {
            max_slide_xml_bytes: default.max_slide_xml_bytes.min(part).min(entry),
            max_total_slide_xml_bytes: default
                .max_total_slide_xml_bytes
                .min(total_part)
                .min(aggregate),
            max_trace_count: default.max_trace_count,
            max_point_count: default.max_point_count,
            max_xml_nodes: default.max_xml_nodes.min(source.max_events()),
            max_xml_depth: default.max_xml_depth.min(source.max_depth()),
            max_output_bytes: default.max_output_bytes.min(part).min(entry),
        }
    }

    /// Intersect this caller-selected semantic policy with the physical OPC
    /// owner policy before a package edit reads, writes, or validates a slide.
    pub(crate) fn constrained_by_source(self, source: litchi_opc::SourceLimits) -> Self {
        let physical = Self::from_source_limits(source);
        Self {
            max_slide_xml_bytes: self.max_slide_xml_bytes.min(physical.max_slide_xml_bytes),
            max_total_slide_xml_bytes: self
                .max_total_slide_xml_bytes
                .min(physical.max_total_slide_xml_bytes),
            max_trace_count: self.max_trace_count.min(physical.max_trace_count),
            max_point_count: self.max_point_count.min(physical.max_point_count),
            max_xml_nodes: self.max_xml_nodes.min(physical.max_xml_nodes),
            max_xml_depth: self.max_xml_depth.min(physical.max_xml_depth),
            max_output_bytes: self.max_output_bytes.min(physical.max_output_bytes),
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Mutable aggregate usage for one or more laser-trace slide reads.
///
/// Pass one value through the complete scope selected by the caller to
/// preserve byte, trace, and point ceilings. A package-wide walk passes all
/// slides; a source-backed batch passes every selected slide, including an
/// unchanged recipient snapshot whose captured source usage avoids a second
/// XML scan. A failed read does not commit its staged charges.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct Usage {
    total_slide_xml_bytes: usize,
    trace_count: usize,
    point_count: usize,
}

impl Usage {
    /// Start an empty aggregate ledger.
    pub const fn new() -> Self {
        Self {
            total_slide_xml_bytes: 0,
            trace_count: 0,
            point_count: 0,
        }
    }

    /// Aggregate slide XML bytes charged by successful reads.
    #[must_use]
    pub const fn total_slide_xml_bytes(self) -> usize {
        self.total_slide_xml_bytes
    }

    /// Aggregate traces charged by successful reads.
    #[must_use]
    pub const fn trace_count(self) -> usize {
        self.trace_count
    }

    /// Aggregate points charged by successful reads.
    #[must_use]
    pub const fn point_count(self) -> usize {
        self.point_count
    }

    pub(super) fn add_slide_xml_bytes(&mut self, bytes: usize) -> Option<usize> {
        let total = self.total_slide_xml_bytes.checked_add(bytes)?;
        self.total_slide_xml_bytes = total;
        Some(total)
    }

    pub(super) fn add_trace(&mut self) -> Option<usize> {
        let total = self.trace_count.checked_add(1)?;
        self.trace_count = total;
        Some(total)
    }

    pub(super) fn add_point(&mut self) -> Option<usize> {
        let total = self.point_count.checked_add(1)?;
        self.point_count = total;
        Some(total)
    }

    /// Add one independently captured slide ledger without changing either
    /// input.  The caller applies its selected policy to the returned total;
    /// this keeps aggregate scope explicit when a batch preserves unchanged
    /// slide snapshots without rescanning their XML.
    pub(crate) const fn checked_add(self, other: Self) -> Option<Self> {
        let total_slide_xml_bytes = match self
            .total_slide_xml_bytes
            .checked_add(other.total_slide_xml_bytes)
        {
            Some(value) => value,
            None => return None,
        };
        let trace_count = match self.trace_count.checked_add(other.trace_count) {
            Some(value) => value,
            None => return None,
        };
        let point_count = match self.point_count.checked_add(other.point_count) {
            Some(value) => value,
            None => return None,
        };
        Some(Self {
            total_slide_xml_bytes,
            trace_count,
            point_count,
        })
    }
}
