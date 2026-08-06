//! Typed ODP embedded-chart values and bounded resource budgets.

use litchi_odf_common::chart::Element;
use litchi_odf_common::drawing::Frame;
use std::sync::Arc;

/// Limits applied while scanning and editing embedded chart parts.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_charts: usize,
    max_part_bytes: usize,
    max_total_bytes: usize,
}

impl Limits {
    /// Construct a positive chart resource budget.
    pub fn new(
        max_charts: usize,
        max_part_bytes: usize,
        max_total_bytes: usize,
    ) -> litchi_core::Result<Self> {
        if max_charts == 0 || max_part_bytes == 0 || max_total_bytes == 0 {
            return Err(litchi_core::Error::InvalidFormat(
                "ODP chart limits must be positive".to_string(),
            ));
        }
        if max_total_bytes < max_part_bytes {
            return Err(litchi_core::Error::InvalidFormat(
                "ODP chart total-byte limit must cover one chart part".to_string(),
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

/// How an ODP drawing occurrence stores its chart content.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
#[non_exhaustive]
pub enum Storage {
    /// A referenced `Object_N/` chart subdocument.
    #[default]
    PackageSubdocument,
    /// An inline `office:document` child of `draw:object`.
    InlineXml,
}

/// A bounded, typed view of one chart `content.xml` part.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Part {
    pub(crate) xml: Arc<str>,
    pub(crate) chart: Arc<Element>,
}

impl Part {
    /// Borrow the exact source or authored chart XML.
    #[must_use]
    pub fn xml(&self) -> &str {
        &self.xml
    }

    /// Borrow the namespace-aware retained chart tree.
    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.chart
    }

    /// Borrow the typed plot-area view when the chart declares one.
    #[must_use]
    pub fn plot_area(&self) -> Option<litchi_odf_common::chart::PlotArea<'_>> {
        self.chart.plot_area()
    }

    /// Borrow the typed legend view when the chart declares one.
    #[must_use]
    pub fn legend(&self) -> Option<litchi_odf_common::chart::Legend<'_>> {
        self.chart.legend()
    }
}

/// One embedded chart occurrence in an ODP drawing page.
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

    /// Borrow the containing frame context.
    #[must_use]
    pub fn frame(&self) -> Option<&Frame> {
        self.frame.as_ref()
    }

    /// Return the exact producer-visible drawing name, if present.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.frame.as_ref().and_then(|frame| frame.name.as_deref())
    }

    /// Return the containing presentation page name, if present.
    #[must_use]
    pub fn page(&self) -> Option<&str> {
        self.frame
            .as_ref()
            .and_then(|frame| frame.page_name.as_deref())
    }

    /// Return the storage form without exposing package paths.
    #[must_use]
    pub const fn storage(&self) -> Storage {
        self.storage
    }

    /// Borrow the exact standalone chart XML.
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
                Location::Existing {
                    content_path: left, ..
                },
                Location::Existing {
                    content_path: right,
                    ..
                },
            ) => left.is_some() && left == right,
            _ => false,
        }
    }

    pub(crate) fn same_identity(&self, other: &Self) -> bool {
        match (&self.location, &other.location) {
            (
                Location::Existing {
                    object_start: left,
                    object_end: left_end,
                    ..
                },
                Location::Existing {
                    object_start: right,
                    object_end: right_end,
                    ..
                },
            ) => left == right && left_end == right_end,
            (Location::Added { token: left, .. }, Location::Added { token: right, .. }) => {
                left == right
            },
            _ => false,
        }
    }
}

/// Selector used by chart inventory and transactions.
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

/// Selector for a presentation page used by chart insertion.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Page<'a> {
    /// Checked zero-based page order.
    Index(usize),
    /// Exact `draw:name` page name.
    Name(&'a str),
}

impl From<usize> for Page<'static> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Page<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

/// Physical ownership used internally to preserve package structure.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum Location {
    Existing {
        object_start: usize,
        object_end: usize,
        payload: Option<(usize, usize)>,
        content_path: Option<String>,
    },
    Added {
        page_index: usize,
        token: usize,
    },
}
