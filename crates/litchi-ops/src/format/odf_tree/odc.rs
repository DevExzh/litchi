//! OpenDocument charts (`.odc`): the chart, its axes, series and legend,
//! read from the `litchi_odc::Chart` element tree and written through
//! `litchi_odc::Edit`.

use litchi_core::FileFormat;
use litchi_core::selector::Path;
use litchi_odc::chart::{Element, Kind as ChartKind};

use super::tree::{count, length, text};
use super::write::{
    Commit, Patch, bound, element_refusal, finish, key_refusal, refusal, resolve, staging_error,
};
use super::{Flavor, OdfTree, Owner, TreeBuilder, element, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::{LengthSpec, ValueType};

kinds! {
    /// The element kinds of an ODC chart.
    Kind: &crate::path::vocabulary::CHART, root Chart {
        Chart => "chart",
        Axis => "axis",
        Series => "series",
        Legend => "legend",
    }
}

const CHART_PROPS: &[PropSpec] = &[
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("subtitle", ValueType::Text),
    PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("axisCount", ValueType::Int),
    PropSpec::new("seriesCount", ValueType::Int),
    PropSpec::new("dataRange", ValueType::Text),
];

const AXIS_PROPS: &[PropSpec] = &[
    PropSpec::new("dimension", ValueType::Text),
    writable("name", ValueType::Text),
    PropSpec::new("title", ValueType::Text),
    PropSpec::new("categories", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
];

const SERIES_PROPS: &[PropSpec] = &[
    PropSpec::new("values", ValueType::Text),
    PropSpec::new("label", ValueType::Text),
    PropSpec::new("chartType", ValueType::Text),
    PropSpec::new("axis", ValueType::Text),
    PropSpec::new("pointCount", ValueType::Int),
];

const LEGEND_PROPS: &[PropSpec] = &[PropSpec::new("position", ValueType::Text)];

static OPERATION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "chart",
            context: "the title, subtitle, type and size of the chart",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODC chart definition writes",
        reason: "the leaf rewrites those through a whole typed definition replacement, which the unified adapter does not author yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "add",
            element: "series",
            context: "a new series or axis",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODC typed series and axis creation",
        reason: "creating a series or an axis needs typed range and class values the unified registry does not have yet",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "series",
            context: "series ranges, labels and classes",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0282 ODC series writes",
        reason: "the leaf replaces a series as a whole typed specification",
    },
];

/// The ODC spec table.
pub(crate) static SPEC: FormatSpec<Kind> = FormatSpec {
    root: Kind::Chart,
    elements: &[
        element(
            Kind::Chart,
            "chart",
            &["chart"],
            &[],
            &[Kind::Axis, Kind::Series, Kind::Legend],
            &["/"],
            CHART_PROPS,
        ),
        element(
            Kind::Axis,
            "axis",
            &["axis"],
            &[Kind::Chart],
            &[],
            &["/axis[N]"],
            AXIS_PROPS,
        ),
        element(
            Kind::Series,
            "series",
            &["series"],
            &[Kind::Chart],
            &[],
            &["/series[N]"],
            SERIES_PROPS,
        ),
        element(
            Kind::Legend,
            "legend",
            &["legend"],
            &[Kind::Chart],
            &[],
            &["/legend"],
            LEGEND_PROPS,
        ),
    ],
    gaps: &[],
};

crate::node::spec::check_format_spec!(SPEC);

/// The ODC flavor.
pub(crate) struct Odc;

impl Flavor for Odc {
    type Kind = Kind;
    const FORMAT: FileFormat = FileFormat::Odc;
    const SPEC: &'static FormatSpec<Kind> = &SPEC;
    const TOP: &'static [Kind] = &[Kind::Chart];
    const SINGLETONS: &'static [Kind] = &[Kind::Legend];
}

/// The elements `set` writes, `add` creates and `remove` deletes.
pub(crate) static WRITES: super::Writes = super::Writes {
    settable: &["axis"],
    addable: &[],
    removable: &["axis", "series"],
    gaps: &OPERATION_GAPS,
};

/// The attribute `local` of `element`, in any namespace.
fn attribute<'a>(element: &'a Element, local: &str) -> Option<&'a str> {
    element
        .attributes()
        .iter()
        .find(|attribute| attribute.local_name() == local)
        .map(litchi_odc::chart::Attribute::value)
}

/// The text of the first `kind` child of `element`.
fn child_text(element: &Element, kind: ChartKind) -> Option<String> {
    element.children_of_kind(kind).next().map(Element::all_text)
}

/// Build the tree of `package`.
pub(crate) fn build(package: &litchi_odc::Package) -> Result<super::Tree> {
    let chart = package.chart();
    let root = chart.chart();
    let plot = chart.plot_area();
    let axes: Vec<_> = plot.map(|plot| plot.axes().collect()).unwrap_or_default();
    let series: Vec<_> = plot.map(|plot| plot.series().collect()).unwrap_or_default();
    let mut tree = TreeBuilder::new(SPEC.elements, Odc::SINGLETONS);
    let title = child_text(root, ChartKind::Title);
    tree.root()
        .prop("chartType", text(attribute(root, "class")))
        .prop("title", text(title.as_deref()))
        .prop(
            "subtitle",
            text(child_text(root, ChartKind::Subtitle).as_deref()),
        )
        .prop("width", length(attribute(root, "width")))
        .prop("height", length(attribute(root, "height")))
        .prop("axisCount", count(axes.len()))
        .prop("seriesCount", count(series.len()))
        .prop(
            "dataRange",
            text(plot.and_then(|plot| plot.cell_range_address())),
        )
        .text(title.unwrap_or_default());
    for (index, axis) in axes.iter().enumerate() {
        let dimension = match axis.dimension() {
            Ok(litchi_odc::chart::Dimension::X) => Some("x"),
            Ok(litchi_odc::chart::Dimension::Y) => Some("y"),
            Ok(litchi_odc::chart::Dimension::Z) => Some("z"),
            _ => None,
        };
        let title = child_text(axis.element(), ChartKind::Title);
        let draft = tree
            .push(super::tree::ROOT, Kind::Axis)
            .prop("dimension", text(dimension))
            .prop("name", text(axis.name()))
            .prop("title", text(title.as_deref()))
            .prop("categories", text(axis.categories_range()))
            .prop("style", text(axis.style_name()))
            .label(axis.name())
            .leaf(index, 0);
        match title {
            Some(title) => draft.text(title).done(),
            None => draft.done(),
        };
    }
    for (index, series) in series.iter().enumerate() {
        tree.push(super::tree::ROOT, Kind::Series)
            .prop("values", text(series.values_range()))
            .prop("label", text(series.label_cell()))
            .prop("chartType", text(series.class_name()))
            .prop("axis", text(series.attached_axis()))
            .prop("pointCount", count(series.data_points().count()))
            .leaf(index, 0)
            .done();
    }
    if let Some(legend) = chart.legend() {
        tree.push(super::tree::ROOT, Kind::Legend)
            .prop(
                "position",
                text(Some(
                    attribute(legend.element(), "legend-position").unwrap_or("end"),
                )),
            )
            .done();
    }
    Ok(tree.build())
}

/// Set axis names.
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odc::Package,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Odc>(file, target, limits.max_expanded_targets())?;
    let mut edit = package.edit();
    let mut operations = 0usize;
    for node in &nodes {
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            match (node.item().type_name(), key) {
                ("axis", "name") => {
                    let name = super::write::text(value, "name")?;
                    if name.as_deref() == node.item().label() {
                        continue;
                    }
                    let update = litchi_odc::AxisUpdate {
                        name: Some(name),
                        style_name: None,
                    };
                    edit.update_axis(node.item().leaf()[0], update)
                        .map_err(|error| staging_error(error, FileFormat::Odc))?;
                },
                ("axis" | "chart" | "series" | "legend", other) => {
                    return Err(key_refusal(node, other));
                },
                _ => return Err(element_refusal("set", node)),
            }
        }
    }
    commit(file, edit, operations, limits)
}

fn commit(
    file: &OdfTree,
    edit: litchi_odc::Edit<'_>,
    operations: usize,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let commit = edit
        .commit()
        .map_err(|error| super::write::commit_error(error, FileFormat::Odc))?;
    let changed = commit.diagnostics().changed();
    let (package, patch) = commit.into_parts();
    finish(
        file,
        changed,
        Owner::Odc(package),
        Patch::Odc(patch),
        operations,
        limits,
    )
}

/// Adding chart elements needs typed specifications.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odc::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let _ = (file, package, parent, limits);
    Err(refusal(
        Code::UnsupportedCapability,
        format!(
            "adding `{}` elements to an ODC chart is not supported",
            node.kind()
        ),
    ))
}

/// Remove axes or series.
pub(crate) fn remove(
    file: &OdfTree,
    package: &litchi_odc::Package,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Commit> {
    let nodes = resolve::<Odc>(file, target, limits.max_expanded_targets())?;
    bound(LimitKind::Operations, nodes.len(), limits.max_operations())?;
    // Remove from the highest position down so earlier removals do not
    // shift later positions.
    let mut axes = Vec::new();
    let mut series = Vec::new();
    for node in &nodes {
        match node.item().type_name() {
            "axis" => axes.push(node.item().leaf()[0]),
            "series" => series.push(node.item().leaf()[0]),
            _ => return Err(element_refusal("remove", node)),
        }
    }
    axes.sort_unstable_by(|left, right| right.cmp(left));
    series.sort_unstable_by(|left, right| right.cmp(left));
    let mut edit = package.edit();
    let staging = |error| staging_error(error, FileFormat::Odc);
    for index in series {
        edit.remove_series(index).map_err(staging)?;
    }
    for index in axes {
        edit.remove_axis(index).map_err(staging)?;
    }
    commit(file, edit, nodes.len(), limits)
}
