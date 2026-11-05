//! The chart bounds one ODC owner reads, edits and replays under.
//!
//! Callers select them through [`litchi_odf_common::OpenOptions`]: the source
//! `input_bytes` bound a package or flat document, and the semantic
//! `content_bytes`, `depth`, `axes`, `series`, `data_points`, `cached_rows`,
//! `cached_cells`, `range_items`, `resources`, `history` and `scalar_bytes`
//! dimensions bound the chart model. The owners keep the options they were
//! opened (or built) with.

use litchi_core::Error;
use litchi_odf_common::{OpenOptions, SemanticLimitKind, compact_xml};

/// The resolved bounds of one chart owner.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Limits {
    package_bytes: usize,
    content_bytes: usize,
    depth: usize,
    axes: usize,
    series: usize,
    data_points: usize,
    cached_rows: usize,
    cached_cells: usize,
    range_items: usize,
    resources: usize,
    history: usize,
    scalar_bytes: usize,
    options: OpenOptions,
}

impl Default for Limits {
    fn default() -> Self {
        Self::from_options(OpenOptions::new())
    }
}

impl Limits {
    /// Resolve the chart bounds of one [`OpenOptions`] envelope. Chart XML
    /// is validated by the compact-XML checker, so the content and depth
    /// bounds never exceed its ceilings.
    pub(crate) fn from_options(options: OpenOptions) -> Self {
        let semantic = options.semantic_limits_value();
        let value = |kind| semantic.usize_value(kind);
        Self {
            package_bytes: usize::try_from(options.source_limits_value().max_input_bytes())
                .unwrap_or(usize::MAX),
            content_bytes: value(SemanticLimitKind::ContentBytes).min(compact_xml::HARD_MAX_BYTES),
            depth: value(SemanticLimitKind::Depth).min(compact_xml::HARD_MAX_DEPTH),
            axes: value(SemanticLimitKind::Axes),
            series: value(SemanticLimitKind::Series),
            data_points: value(SemanticLimitKind::DataPoints),
            cached_rows: value(SemanticLimitKind::CachedRows),
            cached_cells: value(SemanticLimitKind::CachedCells),
            range_items: value(SemanticLimitKind::RangeItems),
            resources: value(SemanticLimitKind::Resources),
            history: value(SemanticLimitKind::History),
            scalar_bytes: value(SemanticLimitKind::ScalarBytes),
            options,
        }
    }

    /// The options these bounds were resolved from.
    pub(crate) const fn options(self) -> OpenOptions {
        self.options
    }

    /// The typed failure of an oversized package or flat document.
    pub(crate) fn input_exceeded(self, observed: usize) -> Error {
        Error::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(self.package_bytes).unwrap_or(u64::MAX),
            scope: "ODF owned package input".into(),
        })
    }

    pub(crate) const fn max_package_bytes(self) -> usize {
        self.package_bytes
    }

    pub(crate) const fn max_content_bytes(self) -> usize {
        self.content_bytes
    }

    pub(crate) const fn max_depth(self) -> usize {
        self.depth
    }

    pub(crate) const fn max_axes(self) -> usize {
        self.axes
    }

    pub(crate) const fn max_series(self) -> usize {
        self.series
    }

    pub(crate) const fn max_data_points(self) -> usize {
        self.data_points
    }

    pub(crate) const fn max_cached_rows(self) -> usize {
        self.cached_rows
    }

    pub(crate) const fn max_cached_cells(self) -> usize {
        self.cached_cells
    }

    pub(crate) const fn max_range_items(self) -> usize {
        self.range_items
    }

    pub(crate) const fn max_resources(self) -> usize {
        self.resources
    }

    pub(crate) const fn max_history(self) -> usize {
        self.history
    }

    pub(crate) const fn max_scalar_bytes(self) -> usize {
        self.scalar_bytes
    }
}
