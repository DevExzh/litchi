//! Semantic laser-trace values and bounded resource state.

use crate::time::Offset;
use litchi_drawingml::coordinate::Coordinate;

pub(super) const PRESENTATIONML_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/presentationml/2006/main";
pub(super) const STRICT_PRESENTATIONML_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/presentationml/main";

/// The PresentationML namespace dialect used by a laser-trace writer.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum Conformance {
    /// The ISO Transitional PresentationML namespace.
    #[default]
    Transitional,
    /// The ISO Strict PresentationML namespace.
    Strict,
}

impl Conformance {
    /// Return the PresentationML namespace URI used by this profile.
    #[inline]
    pub const fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => PRESENTATIONML_NAMESPACE,
            Self::Strict => STRICT_PRESENTATIONML_NAMESPACE,
        }
    }

    /// Select the profile for a detected PresentationML namespace.
    #[inline]
    pub fn from_namespace(namespace: &str) -> Self {
        if namespace == STRICT_PRESENTATIONML_NAMESPACE {
            Self::Strict
        } else {
            Self::Transitional
        }
    }
}

/// A persisted laser-pointer point from a PowerPoint slide show.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TracePoint {
    pub(super) time: Offset,
    pub(super) x: Coordinate,
    pub(super) y: Coordinate,
}

impl TracePoint {
    /// Create a trace point from exact, checked time and coordinate values.
    pub fn new(time: Offset, x: Coordinate, y: Coordinate) -> Self {
        Self { time, x, y }
    }

    /// Return the exact normalized time offset relative to the slide timeline.
    #[inline]
    pub fn time(&self) -> &Offset {
        &self.time
    }

    /// Return the checked horizontal DrawingML coordinate.
    #[inline]
    pub fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Return the checked vertical DrawingML coordinate.
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
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of this trace on its slide.
    #[inline]
    pub fn trace_index(&self) -> usize {
        self.trace_index
    }

    /// Return the stored trace points in source order.
    #[inline]
    pub fn points(&self) -> &[TracePoint] {
        &self.points
    }

    /// Return the number of stored trace points.
    #[inline]
    pub fn point_count(&self) -> usize {
        self.points.len()
    }
}

/// Aggregate resource state for reading one or more slide XML parts.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Limits {
    pub(super) total_slide_xml_bytes: usize,
    pub(super) trace_count: usize,
    pub(super) point_count: usize,
}
