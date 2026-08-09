//! Typed ODS embedded-chart values and resource limits.

use litchi_odf_common::chart::Element;
use litchi_odf_common::drawing::Frame;
use std::sync::Arc;

/// Bounds applied while discovering and replacing embedded chart parts.
///
/// The limits are deliberately part of the public chart owner: callers that
/// process untrusted workbooks can choose a smaller budget without changing
/// the semantic API or exposing archive internals.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_charts: usize,
    max_part_bytes: usize,
    max_total_bytes: usize,
}

impl Limits {
    /// Create a checked chart resource budget.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(
        max_charts: usize,
        max_part_bytes: usize,
        max_total_bytes: usize,
    ) -> litchi_core::Result<Self> {
        if max_charts == 0 || max_part_bytes == 0 || max_total_bytes == 0 {
            return Err(litchi_core::Error::InvalidFormat(
                "ODS chart limits must be positive".to_string(),
            ));
        }
        if max_total_bytes < max_part_bytes {
            return Err(litchi_core::Error::InvalidFormat(
                "ODS chart total-byte limit must cover one chart part".to_string(),
            ));
        }
        Ok(Self {
            max_charts,
            max_part_bytes,
            max_total_bytes,
        })
    }

    #[must_use]
    pub const fn max_charts(self) -> usize {
        self.max_charts
    }

    #[must_use]
    pub const fn max_part_bytes(self) -> usize {
        self.max_part_bytes
    }

    #[must_use]
    pub const fn max_total_bytes(self) -> usize {
        self.max_total_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_charts: 4_096,
            max_part_bytes: 16 * 1024 * 1024,
            max_total_bytes: 64 * 1024 * 1024,
        }
    }
}

/// How an ODS drawing occurrence stores its chart content.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Storage {
    /// A package subdocument such as `Object_1/content.xml`.
    PackageSubdocument,
    /// An inline `office:document` child of `draw:object`.
    InlineXml,
}

/// A validated standalone chart content part.
///
/// The source XML is retained exactly for package-backed parts. The parsed
/// common [`Element`] tree is a second, typed view; unknown namespaces and
/// attributes remain in that tree and are never discarded by this owner.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Part {
    pub(crate) xml: Arc<str>,
    pub(crate) chart: Arc<Element>,
}

impl Part {
    /// Borrow the validated standalone chart `content.xml` representation.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Borrow the namespace-aware retained chart tree.
    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.chart
    }

    /// Borrow the first typed plot-area view, if present.
    #[must_use]
    pub fn plot_area(&self) -> Option<litchi_odf_common::chart::PlotArea<'_>> {
        self.chart.plot_area()
    }

    /// Borrow the first typed legend view, if present.
    #[must_use]
    pub fn legend(&self) -> Option<litchi_odf_common::chart::Legend<'_>> {
        self.chart.legend()
    }
}

/// One embedded chart occurrence in the ODS drawing graph.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Chart {
    pub(crate) frame: Option<Frame>,
    pub(crate) storage: Storage,
    pub(crate) part: Part,
    pub(crate) location: Location,
}

impl Chart {
    /// Borrow the chart content part.
    #[must_use]
    pub fn part(&self) -> &Part {
        &self.part
    }

    /// Borrow the drawing frame context, if the object is framed.
    #[must_use]
    pub fn frame(&self) -> Option<&Frame> {
        self.frame.as_ref()
    }

    /// Return the producer-visible drawing name, if present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.frame.as_ref().and_then(|frame| frame.name.as_deref())
    }

    /// Return the containing spreadsheet sheet name, if present.
    #[must_use]
    pub fn sheet(&self) -> Option<&str> {
        self.frame
            .as_ref()
            .and_then(|frame| frame.sheet_name.as_deref())
    }

    /// Return the storage form without exposing an archive path.
    #[must_use]
    pub const fn storage(&self) -> Storage {
        self.storage
    }

    /// Borrow the standalone chart XML without copying it.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.part.xml()
    }

    pub(crate) fn with_part(&self, part: Part) -> Self {
        Self {
            frame: self.frame.clone(),
            storage: self.storage,
            part,
            location: self.location.clone(),
        }
    }

    pub(crate) fn shares_storage(&self, other: &Self) -> bool {
        match (&self.location, &other.location) {
            (
                Location::Package { content_path: left },
                Location::Package {
                    content_path: right,
                },
            ) => left == right,
            _ => false,
        }
    }
}

/// Selector used by inventory and transaction APIs.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Checked zero-based discovery order.
    Index(usize),
    /// Exact producer-visible drawing name.
    Name(&'a str),
}

impl From<usize> for Selector<'static> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

/// Internal package location retained by a chart occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum Location {
    Package {
        content_path: String,
    },
    Inline {
        payload_start: usize,
        payload_end: usize,
    },
}
