//! Non-mutating semantic three-way merge support.

use crate::{Definition, DefinitionPatch, DefinitionSnapshot, Limits, PlotAreaSpec};

/// One deterministic semantic merge conflict.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct Conflict {
    path: String,
}

impl Conflict {
    pub(crate) fn new(path: impl Into<String>) -> Self {
        Self { path: path.into() }
    }

    /// Return the stable semantic path that conflicted.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }
}

/// Result of a non-mutating definition join or dependency transfer.
pub struct DefinitionMerge {
    patch: Option<DefinitionPatch>,
    conflicts: Vec<Conflict>,
}

impl DefinitionMerge {
    pub(crate) fn new(patch: Option<DefinitionPatch>, mut conflicts: Vec<Conflict>) -> Self {
        conflicts.sort();
        conflicts.dedup();
        Self { patch, conflicts }
    }

    /// Return the merged patch when no conflicts occurred.
    #[must_use]
    pub const fn patch(&self) -> Option<&DefinitionPatch> {
        self.patch.as_ref()
    }

    /// Consume the result into the merged patch.
    #[must_use]
    pub fn into_patch(self) -> Option<DefinitionPatch> {
        self.patch
    }

    /// Return deterministic conflicts without mutating any input snapshot.
    #[must_use]
    pub fn conflicts(&self) -> &[Conflict] {
        &self.conflicts
    }

    /// Return whether the merge completed without conflicts.
    #[must_use]
    pub fn is_merged(&self) -> bool {
        self.patch.is_some()
    }
}

pub(crate) fn join(left: &DefinitionPatch, right: &DefinitionPatch) -> DefinitionMerge {
    if left.source.limits != right.source.limits
        || *left.source.definition != *right.source.definition
    {
        return DefinitionMerge::new(None, vec![Conflict::new("chart.source")]);
    }
    merge_patch(
        &left.source,
        &left.target,
        &right.target,
        &left.source,
        left.changes.iter().chain(&right.changes).cloned().collect(),
        left.source.limits,
    )
}

pub(crate) fn transfer(
    patch: &DefinitionPatch,
    destination: &DefinitionSnapshot,
) -> DefinitionMerge {
    merge_patch(
        &patch.source,
        &patch.target,
        destination,
        destination,
        patch.changes.clone(),
        destination.limits,
    )
}

fn merge_patch(
    base: &DefinitionSnapshot,
    left: &DefinitionSnapshot,
    right: &DefinitionSnapshot,
    source: &DefinitionSnapshot,
    changes: Vec<crate::DefinitionChange>,
    limits: Limits,
) -> DefinitionMerge {
    let mut conflicts = Vec::new();
    let candidate = merge_definition(
        &base.definition,
        &left.definition,
        &right.definition,
        &mut conflicts,
    );
    if !conflicts.is_empty() {
        return DefinitionMerge::new(None, conflicts);
    }
    let target = match DefinitionSnapshot::new(candidate, limits) {
        Ok(snapshot) => snapshot,
        Err(_validation_error) => {
            return DefinitionMerge::new(None, vec![Conflict::new("chart.dependencies")]);
        },
    };
    DefinitionMerge::new(
        Some(DefinitionPatch {
            source: source.clone(),
            target,
            changes,
        }),
        Vec::new(),
    )
}

fn merge_definition(
    base: &Definition,
    left: &Definition,
    right: &Definition,
    conflicts: &mut Vec<Conflict>,
) -> Definition {
    Definition {
        class: merge_value(
            "chart.class",
            &base.class,
            &left.class,
            &right.class,
            conflicts,
        ),
        style_name: merge_value(
            "chart.style",
            &base.style_name,
            &left.style_name,
            &right.style_name,
            conflicts,
        ),
        width: merge_value(
            "chart.width",
            &base.width,
            &left.width,
            &right.width,
            conflicts,
        ),
        height: merge_value(
            "chart.height",
            &base.height,
            &left.height,
            &right.height,
            conflicts,
        ),
        title: merge_value(
            "chart.title",
            &base.title,
            &left.title,
            &right.title,
            conflicts,
        ),
        subtitle: merge_value(
            "chart.subtitle",
            &base.subtitle,
            &left.subtitle,
            &right.subtitle,
            conflicts,
        ),
        footer: merge_value(
            "chart.footer",
            &base.footer,
            &left.footer,
            &right.footer,
            conflicts,
        ),
        legend: merge_value(
            "chart.legend",
            &base.legend,
            &left.legend,
            &right.legend,
            conflicts,
        ),
        plot_area: merge_plot(
            &base.plot_area,
            &left.plot_area,
            &right.plot_area,
            conflicts,
        ),
        cached_table: merge_value(
            "chart.data",
            &base.cached_table,
            &left.cached_table,
            &right.cached_table,
            conflicts,
        ),
        calculation_settings: merge_value(
            "chart.calculation-settings",
            &base.calculation_settings,
            &left.calculation_settings,
            &right.calculation_settings,
            conflicts,
        ),
        extensions: merge_value(
            "chart.extensions",
            &base.extensions,
            &left.extensions,
            &right.extensions,
            conflicts,
        ),
    }
}

fn merge_plot(
    base: &PlotAreaSpec,
    left: &PlotAreaSpec,
    right: &PlotAreaSpec,
    conflicts: &mut Vec<Conflict>,
) -> PlotAreaSpec {
    PlotAreaSpec {
        cell_range_address: merge_value(
            "chart.plot.range",
            &base.cell_range_address,
            &left.cell_range_address,
            &right.cell_range_address,
            conflicts,
        ),
        data_source_labels: merge_value(
            "chart.plot.labels",
            &base.data_source_labels,
            &left.data_source_labels,
            &right.data_source_labels,
            conflicts,
        ),
        style_name: merge_value(
            "chart.plot.style",
            &base.style_name,
            &left.style_name,
            &right.style_name,
            conflicts,
        ),
        x: merge_value("chart.plot.x", &base.x, &left.x, &right.x, conflicts),
        y: merge_value("chart.plot.y", &base.y, &left.y, &right.y, conflicts),
        width: merge_value(
            "chart.plot.width",
            &base.width,
            &left.width,
            &right.width,
            conflicts,
        ),
        height: merge_value(
            "chart.plot.height",
            &base.height,
            &left.height,
            &right.height,
            conflicts,
        ),
        axes: merge_vector(
            "chart.plot.axes",
            &base.axes,
            &left.axes,
            &right.axes,
            conflicts,
        ),
        series: merge_vector(
            "chart.plot.series",
            &base.series,
            &left.series,
            &right.series,
            conflicts,
        ),
        wall: merge_value(
            "chart.plot.wall",
            &base.wall,
            &left.wall,
            &right.wall,
            conflicts,
        ),
        floor: merge_value(
            "chart.plot.floor",
            &base.floor,
            &left.floor,
            &right.floor,
            conflicts,
        ),
        stock_gain_marker: merge_value(
            "chart.plot.stock-gain-marker",
            &base.stock_gain_marker,
            &left.stock_gain_marker,
            &right.stock_gain_marker,
            conflicts,
        ),
        stock_loss_marker: merge_value(
            "chart.plot.stock-loss-marker",
            &base.stock_loss_marker,
            &left.stock_loss_marker,
            &right.stock_loss_marker,
            conflicts,
        ),
        stock_range_line: merge_value(
            "chart.plot.stock-range-line",
            &base.stock_range_line,
            &left.stock_range_line,
            &right.stock_range_line,
            conflicts,
        ),
        extensions: merge_value(
            "chart.plot.extensions",
            &base.extensions,
            &left.extensions,
            &right.extensions,
            conflicts,
        ),
    }
}

fn merge_vector<T: Clone + PartialEq>(
    path: &str,
    base: &[T],
    left: &[T],
    right: &[T],
    conflicts: &mut Vec<Conflict>,
) -> Vec<T> {
    if base.len() == left.len() && base.len() == right.len() {
        return base
            .iter()
            .zip(left)
            .zip(right)
            .enumerate()
            .map(|(index, ((base_item, left_item), right_item))| {
                merge_value(
                    &format!("{path}[{index}]"),
                    base_item,
                    left_item,
                    right_item,
                    conflicts,
                )
            })
            .collect();
    }
    merge_value(
        path,
        &base.to_vec(),
        &left.to_vec(),
        &right.to_vec(),
        conflicts,
    )
}

fn merge_value<T: Clone + PartialEq>(
    path: &str,
    base: &T,
    left: &T,
    right: &T,
    conflicts: &mut Vec<Conflict>,
) -> T {
    if left == right {
        left.clone()
    } else if left == base {
        right.clone()
    } else if right == base {
        left.clone()
    } else {
        conflicts.push(Conflict::new(path));
        base.clone()
    }
}
