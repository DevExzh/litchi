//! OpenDocument charts (`.odc`): the chart, its axes, series and legend,
//! read from the `litchi_odc::Chart` element tree and written through
//! `litchi_odc::Edit`.

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_odc::chart::{Dimension, Element, Kind as ChartKind};
use litchi_odc::{AxisSpec, ChartClass, Definition, SeriesSpec, Text};

use super::tree::{count, length, odf_length, text};
use super::write::{
    Commit, Patch, added_path, bound, element_refusal, finish, key_refusal, refusal, required,
    resolve, staging_error,
};
use super::{Flavor, OdfTree, Owner, TreeBuilder, addable, element, writable};
use crate::Result;
use crate::error::Code;
use crate::new_node::NewNode;
use crate::node::spec::{FormatSpec, GapSpec, PropSpec, Subject};
use crate::path::PathArg;
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

/// A readable property `add` accepts but `set` does not write.
const fn add_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.add = true;
    spec
}

const CHART_PROPS: &[PropSpec] = &[
    writable("chartType", ValueType::Text),
    writable("title", ValueType::Text),
    writable("subtitle", ValueType::Text),
    writable("width", ValueType::Length(LengthSpec::EXTENT)),
    writable("height", ValueType::Length(LengthSpec::EXTENT)),
    PropSpec::new("axisCount", ValueType::Int),
    PropSpec::new("seriesCount", ValueType::Int),
    PropSpec::new("dataRange", ValueType::Text),
];

const AXIS_PROPS: &[PropSpec] = &[
    add_only("dimension", ValueType::Text),
    addable("name", ValueType::Text),
    addable("title", ValueType::Text),
    addable("categories", ValueType::Text),
    PropSpec::new("style", ValueType::Text),
];

const SERIES_PROPS: &[PropSpec] = &[
    addable("values", ValueType::Text),
    addable("label", ValueType::Text),
    addable("chartType", ValueType::Text),
    addable("axis", ValueType::Text),
    PropSpec::new("pointCount", ValueType::Int),
];

const LEGEND_PROPS: &[PropSpec] = &[PropSpec::new("position", ValueType::Text)];

static OPERATION_GAPS: [GapSpec; 3] = [
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "chart",
            context: "a chart outside the leaf's lossless typed definition surface",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0288 ODC lossless producer definitions",
        reason: "title, type, size, series and axis writes replace the chart through its typed definition; a producer chart the leaf cannot project losslessly is refused rather than rewritten",
    },
    GapSpec {
        subject: Subject::Operation {
            verb: "set",
            element: "series",
            context: "data points, data labels, regressions and series styles",
        },
        code: Code::UnsupportedCapability,
        follow_up: "0288 ODC series detail writes",
        reason: "the unified registry writes a series' ranges, label, class and axis; its data points, labels, regression curves and styles stay as authored",
    },
    super::METADATA_PATH_GAP,
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
    settable: &["chart", "axis", "series"],
    addable: &["axis", "series"],
    removable: &["axis", "series"],
    movable: &[],
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
            Ok(Dimension::X) => Some("x"),
            Ok(Dimension::Y) => Some("y"),
            Ok(Dimension::Z) => Some("z"),
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

/// Set chart, axis and series properties. Axis renames alone stay exact
/// attribute splices; any other write replaces the chart through its typed
/// definition (lossless-or-refuse).
pub(crate) fn set(
    file: &OdfTree,
    package: &litchi_odc::Package,
    groups: &[super::write::Group<'_, '_, '_>],
    limits: &SemanticLimits,
) -> Result<Commit> {
    let mut writes = Vec::new();
    let mut operations = 0usize;
    for (node, props) in groups
        .iter()
        .flat_map(|(nodes, props)| nodes.iter().map(move |node| (node, *props)))
    {
        for (key, value) in props.iter() {
            let key = key.as_str();
            operations += 1;
            bound(LimitKind::Operations, operations, limits.max_operations())?;
            let index = node.item().leaf()[0];
            let write = match (node.item().type_name(), key) {
                ("chart", "title") => Write::Title(super::write::text(value, "title")?),
                ("chart", "subtitle") => Write::Subtitle(super::write::text(value, "subtitle")?),
                ("chart", "chartType") => Write::Class(chart_class(&required(
                    super::write::text(value, "chartType")?,
                    "chartType",
                )?)?),
                ("chart", key @ ("width" | "height")) => {
                    let name = if key == "width" { "width" } else { "height" };
                    let length = super::write::length(value, name, LengthSpec::EXTENT)?;
                    Write::Size(name == "width", length.map(odf_length))
                },
                ("axis", "name") => {
                    let name = super::write::text(value, "name")?;
                    if name.as_deref() == node.item().label() {
                        continue;
                    }
                    Write::AxisName(index, name)
                },
                ("axis", "title") => Write::AxisTitle(index, super::write::text(value, "title")?),
                ("axis", "categories") => {
                    Write::AxisCategories(index, super::write::text(value, "categories")?)
                },
                ("series", "values") => Write::Series(
                    index,
                    SeriesField::Values,
                    super::write::text(value, "values")?,
                ),
                ("series", "label") => Write::Series(
                    index,
                    SeriesField::Label,
                    super::write::text(value, "label")?,
                ),
                ("series", "axis") => {
                    Write::Series(index, SeriesField::Axis, super::write::text(value, "axis")?)
                },
                ("series", "chartType") => {
                    let class = super::write::text(value, "chartType")?
                        .map(|value| chart_class(&value))
                        .transpose()?;
                    Write::SeriesClass(index, class)
                },
                ("axis" | "chart" | "series" | "legend", other) => {
                    return Err(key_refusal(node, other));
                },
                _ => return Err(element_refusal("set", node)),
            };
            writes.push(write);
        }
    }
    // A chart the leaf projects losslessly (litchi-authored) is rewritten
    // through its typed definition, so it stays projectable. A producer
    // chart keeps attribute writes as exact-span splices; its title text
    // has no exact form. Axis renames alone are always exact splices.
    let renames_only = writes
        .iter()
        .all(|write| matches!(write, Write::AxisName(..)));
    let typed = if renames_only {
        None
    } else {
        Some(definition(package))
    };
    let source = match typed {
        Some(Ok(source)) => source,
        Some(Err(error)) if !writes.iter().all(Write::is_exact) => return Err(error),
        _ => {
            let mut edit = package.edit();
            for write in &writes {
                write.stage_exact(&mut edit)?;
            }
            return commit(file, edit, operations, limits);
        },
    };
    let mut edit = package.edit();
    let mut definition = source.clone();
    for write in writes {
        write.apply(&mut definition)?;
    }
    if definition != source {
        edit.replace_chart(&definition)
            .map_err(|error| staging_error(error, FileFormat::Odc))?;
    }
    commit(file, edit, operations, limits)
}

/// One typed chart write.
enum Write {
    Title(Option<String>),
    Subtitle(Option<String>),
    Class(ChartClass),
    Size(bool, Option<String>),
    AxisName(usize, Option<String>),
    AxisTitle(usize, Option<String>),
    AxisCategories(usize, Option<String>),
    Series(usize, SeriesField, Option<String>),
    SeriesClass(usize, Option<ChartClass>),
}

/// A text-valued series field.
#[derive(Clone, Copy)]
enum SeriesField {
    Values,
    Label,
    Axis,
}

impl Write {
    /// Whether the write has an exact-span attribute form.
    const fn is_exact(&self) -> bool {
        !matches!(
            self,
            Self::Title(_) | Self::Subtitle(_) | Self::AxisTitle(..)
        )
    }

    /// Stage the exact-span attribute form of the write.
    fn stage_exact(&self, edit: &mut litchi_odc::Edit<'_>) -> Result<()> {
        use litchi_odc::{ExactAttribute as A, ExactTarget as T};
        let (target, attribute, value) = match self {
            Self::AxisName(index, name) => {
                let update = litchi_odc::AxisUpdate {
                    name: Some(name.clone()),
                    style_name: None,
                };
                return edit
                    .update_axis(*index, update)
                    .map_err(|error| staging_error(error, FileFormat::Odc));
            },
            Self::Class(class) => (T::Chart, A::Class, Some(class.lexical().to_owned())),
            Self::Size(width, value) => (
                T::Chart,
                if *width { A::Width } else { A::Height },
                value.clone(),
            ),
            Self::AxisCategories(axis, value) => (
                T::Categories { axis: *axis },
                A::CellRangeAddress,
                value.clone(),
            ),
            Self::Series(index, field, value) => (
                T::Series(*index),
                match field {
                    SeriesField::Values => A::ValuesCellRangeAddress,
                    SeriesField::Label => A::LabelCellAddress,
                    SeriesField::Axis => A::AttachedAxis,
                },
                value.clone(),
            ),
            Self::SeriesClass(index, class) => (
                T::Series(*index),
                A::Class,
                class.as_ref().map(|class| class.lexical().to_owned()),
            ),
            Self::Title(_) | Self::Subtitle(_) | Self::AxisTitle(..) => {
                return Err(refusal(
                    Code::InternalError,
                    "an ODC text write has no exact-span form",
                ));
            },
        };
        edit.update_exact(target, attribute, value)
            .map_err(|error| staging_error(error, FileFormat::Odc))
    }

    fn apply(self, definition: &mut Definition) -> Result<()> {
        let missing = || refusal(Code::InternalError, "the ODC element disappeared");
        match self {
            Self::Title(value) => definition.title = retitle(definition.title.take(), value),
            Self::Subtitle(value) => {
                definition.subtitle = retitle(definition.subtitle.take(), value);
            },
            Self::Class(class) => definition.class = class,
            Self::Size(width, value) => {
                if width {
                    definition.width = value;
                } else {
                    definition.height = value;
                }
            },
            Self::AxisName(index, value) => {
                definition
                    .plot_area
                    .axes
                    .get_mut(index)
                    .ok_or_else(missing)?
                    .name = value;
            },
            Self::AxisTitle(index, value) => {
                let axis = definition
                    .plot_area
                    .axes
                    .get_mut(index)
                    .ok_or_else(missing)?;
                axis.title = retitle(axis.title.take(), value);
            },
            Self::AxisCategories(index, value) => {
                definition
                    .plot_area
                    .axes
                    .get_mut(index)
                    .ok_or_else(missing)?
                    .categories_cell_range_address = value;
            },
            Self::Series(index, field, value) => {
                let series = definition
                    .plot_area
                    .series
                    .get_mut(index)
                    .ok_or_else(missing)?;
                match field {
                    SeriesField::Values => series.values_cell_range_address = value,
                    SeriesField::Label => series.label_cell_address = value,
                    SeriesField::Axis => series.attached_axis = value,
                }
            },
            Self::SeriesClass(index, class) => {
                definition
                    .plot_area
                    .series
                    .get_mut(index)
                    .ok_or_else(missing)?
                    .class = class;
            },
        }
        Ok(())
    }
}

/// `title` with its text replaced (keeping its style and position), or
/// removed for `null`.
fn retitle(title: Option<Text>, value: Option<String>) -> Option<Text> {
    let value = value?;
    Some(match title {
        Some(mut title) => {
            title.text = value;
            title
        },
        None => Text::new(value),
    })
}

/// The chart's typed definition; a producer chart outside the leaf's
/// lossless typed surface is a capability gap, not a corrupt file.
fn definition(package: &litchi_odc::Package) -> Result<Definition> {
    package.chart().definition().map_err(|error| {
        let mut refused = staging_error(error, FileFormat::Odc);
        if refused.code() == Code::InvalidValue {
            refused = refusal(
                Code::UnsupportedCapability,
                "the chart cannot be projected losslessly into a typed definition",
            )
            .with_format(FileFormat::Odc)
            .with_suggestion("follow-up: 0288 ODC lossless producer definitions");
        }
        refused
    })
}

/// A `chart:class` from a unified chart type (`bar` or `chart:bar`).
fn chart_class(value: &str) -> Result<ChartClass> {
    let local = value.strip_prefix("chart:").unwrap_or(value);
    Ok(match local {
        "area" => ChartClass::area(),
        "bar" => ChartClass::bar(),
        "bubble" => ChartClass::bubble(),
        "circle" | "pie" => ChartClass::circle(),
        "filled-radar" => ChartClass::filled_radar(),
        "gantt" => ChartClass::gantt(),
        "line" => ChartClass::line(),
        "radar" => ChartClass::radar(),
        "ring" | "doughnut" => ChartClass::ring(),
        "scatter" => ChartClass::scatter(),
        "stock" => ChartClass::stock(),
        "surface" => ChartClass::surface(),
        _ => {
            return Err(refusal(Code::InvalidValue, "unknown ODF chart type")
                .with_property("chartType".to_owned())
                .with_suggestion(
                    "use area, bar, bubble, circle, filled-radar, gantt, line, radar, ring, scatter, stock or surface",
                ));
        },
    })
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

/// Add a series or an axis to the chart, through its typed definition.
pub(crate) fn add(
    file: &OdfTree,
    package: &litchi_odc::Package,
    parent: &PathArg<'_>,
    node: &NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(Commit, Path<'static>)> {
    let parent_node = super::write::parent::<Odc>(file, parent)?;
    let kind = node.kind().to_ascii_lowercase();
    if parent_node.item().type_name() != "chart" || !matches!(kind.as_str(), "series" | "axis") {
        return Err(refusal(
            Code::UnsupportedCapability,
            format!(
                "adding `{}` under `{}` is not supported for ODC charts",
                node.kind(),
                parent_node.item().type_name()
            ),
        )
        .with_suggestion("add a series or an axis under `/`"));
    }
    let mut definition = definition(package)?;
    let mut props = std::collections::HashMap::new();
    for (key, value) in node.props().iter() {
        props.insert(key.as_str().to_owned(), value);
    }
    let mut take = |key: &'static str| -> Result<Option<String>> {
        props
            .remove(key)
            .map_or(Ok(None), |value| super::write::text(value, key))
    };
    let position;
    if kind == "series" {
        let mut series = SeriesSpec {
            values_cell_range_address: Some(required(take("values")?, "values")?),
            label_cell_address: take("label")?,
            attached_axis: take("axis")?,
            ..SeriesSpec::default()
        };
        series.class = take("chartType")?
            .map(|value| chart_class(&value))
            .transpose()?;
        position = insertion(node, definition.plot_area.series.len())?;
        definition.plot_area.series.insert(position, series);
    } else {
        let dimension = match required(take("dimension")?, "dimension")?.as_str() {
            "x" => Dimension::X,
            "y" => Dimension::Y,
            "z" => Dimension::Z,
            _ => {
                return Err(
                    refusal(Code::InvalidValue, "an axis dimension is x, y or z")
                        .with_property("dimension".to_owned()),
                );
            },
        };
        let mut axis = AxisSpec::new(dimension);
        axis.name = take("name")?;
        axis.title = take("title")?.map(Text::new);
        axis.categories_cell_range_address = take("categories")?;
        position = insertion(node, definition.plot_area.axes.len())?;
        definition.plot_area.axes.insert(position, axis);
    }
    if let Some(key) = props.keys().next() {
        return Err(refusal(
            Code::UnsupportedProperty,
            format!("`{key}` is not a property `add` accepts for an ODC {kind}"),
        )
        .with_property(key.clone()));
    }
    let mut edit = package.edit();
    edit.replace_chart(&definition)
        .map_err(|error| staging_error(error, FileFormat::Odc))?;
    let commit = commit(file, edit, 1, limits)?;
    let path = added_path(&commit.after, &[], &kind, Some(position))?;
    Ok((commit, path))
}

/// The position an `add` placement names among `count` siblings.
fn insertion(node: &NewNode<'_>, count: usize) -> Result<usize> {
    use crate::Placement;
    let index = match node.placement() {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position) | Placement::Into(_, Some(position))) => {
            Position::get(*position)
        },
        Some(Placement::Into(_, None)) => count,
        Some(Placement::Before(_) | Placement::After(_)) => {
            return Err(refusal(
                Code::UnsupportedCapability,
                "ODC series and axes are placed by position",
            ));
        },
    };
    if index > count {
        return Err(refusal(
            Code::OutOfRange,
            "the ODC insertion position is outside the chart",
        ));
    }
    Ok(index)
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
