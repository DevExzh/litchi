//! Typed chart CRUD on slides (record 0270).
//!
//! - `add(/slide[N], chart, {chartType, title, data, categories, legend,
//!   catTitle, axisTitle, name, x, y, width, height})` builds an authored
//!   chart spec ([`crate::format::chart`]) and stages its chart part, its
//!   deterministic embedded workbook, both relationships and the graphic
//!   frame through `litchi_pptx::opened::Edit::add_chart`.
//! - `set` on `/slide[N]/chart[M]` (title, legend, categories, catTitle),
//!   `/slide[N]/chart[M]/series[K]` (name, values) and
//!   `/slide[N]/chart[M]/axis[@role=ROLE]` (title) runs one
//!   source-preserving chart edit and replaces the chart part in place;
//!   a data edit regenerates the embedded workbook when the chart refers to
//!   the authored layout, and is refused otherwise, so the workbook never
//!   silently diverges from the chart's cached values.
//! - `remove` of a chart frame is the ordinary shape removal, which drops the
//!   chart part and its workbook with the frame.
//!
//! Source owners are promoted to an eager package for these edits, since
//! the source-backed slide edit does not stage new or replaced parts.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};

use super::node::{ChartItem, Item, TreeItem};
use super::spec::{
    AXIS_WRITE_TABLE, AxisSetter, CHART_WRITE_TABLE, ChartSetter, Kind, SERIES_WRITE_TABLE,
    SeriesSetter,
};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options};
use super::{Editable, Nav, Pptx};
use crate::format::chart::{self as shared, AddProps, ChartEdit, Role, Summary};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Props, Result};

/// The default chart frame: 1in from the top-left corner, 6in × 4in.
const DEFAULT_BOUNDS: (i64, i64, i64, i64) = (914_400, 914_400, 5_486_400, 3_657_600);

fn text(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(value) | Scalar::Token(value)) => Some(value.to_owned()),
        _ => None,
    }
}

/// The chart a tree item's frame names, read from its chart part.
pub(super) fn summary(file: &Pptx, tree: &TreeItem) -> Result<Option<Arc<Summary>>> {
    let Some(slide) = tree.container.slide() else {
        return Ok(None);
    };
    let Some(payload) = litchi_pptx::shape::raw::payload(tree.handle.view())
        .map_err(|error| error_text(Code::CorruptFile, error.to_string()))?
    else {
        return Ok(None);
    };
    let relationship_id = payload
        .relationship_id()
        .map_err(|error| error_text(Code::CorruptFile, error.to_string()))?;
    let Some(xml) = file.slide_chart_xml(slide as usize, &relationship_id)? else {
        return Ok(None);
    };
    Summary::read(&xml).map(|summary| Some(Arc::new(summary)))
}

fn error_text(code: Code, message: String) -> crate::Error {
    crate::Error::new(code, message).with_format(litchi_core::FileFormat::Pptx)
}

/// What one admitted chart `set` changes, by target.
enum SetTarget {
    Chart,
    Series(u32),
    Axis(u32),
}

impl Pptx {
    pub(super) fn add_chart(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let parent = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = parent.item() else {
            return Err(error(
                Code::UnsupportedType,
                "charts require a slide parent",
            ));
        };
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "chart insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut props = AddProps::default();
        let mut name = None;
        let mut geometry = [None; 4];
        let mut assigned = shared::Assigned::default();
        let mut staged_bytes = 0_usize;
        for (key, value) in node.props().iter() {
            let (index, entry) = CHART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            if !assigned.insert(index) {
                return Err(error(
                    Code::InvalidArgument,
                    "a chart property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            entry.prepare(Operation::Add, value, |setter, scalar| {
                if let Some(Scalar::Text(value)) = scalar {
                    staged_bytes = staged_bytes.saturating_add(value.len());
                    bound(
                        LimitKind::StagedBytes,
                        staged_bytes,
                        limits.max_staged_bytes(),
                    )?;
                }
                match setter {
                    ChartSetter::X | ChartSetter::Y | ChartSetter::Width | ChartSetter::Height => {
                        let Some(Scalar::Length(value)) = scalar else {
                            unreachable!("chart geometry coercion")
                        };
                        let slot = match setter {
                            ChartSetter::X => 0,
                            ChartSetter::Y => 1,
                            ChartSetter::Width => 2,
                            _ => 3,
                        };
                        geometry[slot] = Some(value.emu());
                    },
                    ChartSetter::Name => {
                        let value = text(scalar).unwrap_or_default();
                        super::write::validate_text(&value)?;
                        name = Some(value);
                    },
                    ChartSetter::ChartType => {
                        props.variant =
                            Some(shared::parse_variant(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::Title => props.title = text(scalar),
                    ChartSetter::Legend => {
                        props.legend =
                            Some(shared::parse_legend(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::Categories => {
                        props.categories = Some(shared::parse_list(
                            &text(scalar).unwrap_or_default(),
                            "categories",
                        )?);
                    },
                    ChartSetter::Data => {
                        props.data = Some(shared::parse_data(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::CatTitle => props.category_axis_title = text(scalar),
                    ChartSetter::AxisTitle => props.value_axis_title = text(scalar),
                    ChartSetter::Format(key) => props.format(key, scalar)?,
                    ChartSetter::Numbered(number) => props.numbered_series(
                        usize::from(number) + 1,
                        &text(scalar).unwrap_or_default(),
                    )?,
                    ChartSetter::Dotted(number, field) => props.dotted_series(
                        usize::from(number) + 1,
                        field,
                        &text(scalar).unwrap_or_default(),
                    )?,
                }
                Ok(())
            })?;
        }
        let spec = props.into_spec()?;
        let bounds = (
            geometry[0].unwrap_or(DEFAULT_BOUNDS.0),
            geometry[1].unwrap_or(DEFAULT_BOUNDS.1),
            geometry[2].unwrap_or(DEFAULT_BOUNDS.2),
            geometry[3].unwrap_or(DEFAULT_BOUNDS.3),
        );
        let slide = slide.position as usize;
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        let ordinal = scene
            .roots()
            .filter(|shape| super::steps::tree_kind(shape) == Some(Kind::Chart))
            .count();
        let name = name.unwrap_or_else(|| format!("Chart {}", ordinal + 1));
        let source = self.editable_owner()?;
        let base = eager(&source, self.0.ooxml_limits)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        edit.add_chart(slide, &spec, &name, bounds)?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("chart").with(Predicate::Index(Position::new(ordinal))),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }

    /// A `set` on a chart, series or axis path; `None` for other targets.
    pub(super) fn try_set_chart(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let Some(last) = path.path().steps().last() else {
            return Ok(None);
        };
        if !["chart", "series", "axis"]
            .iter()
            .any(|name| last.name().eq_ignore_ascii_case(name))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let (tree, which) = match node.item() {
            Item::Tree(tree) if tree.kind == Kind::Chart => (tree.clone(), SetTarget::Chart),
            Item::Series(item) => {
                let index = item
                    .summary
                    .series
                    .get(item.index as usize)
                    .map(|series| series.index)
                    .ok_or_else(|| error(Code::PathNotFound, "the series is missing"))?;
                (item.chart.clone(), SetTarget::Series(index))
            },
            Item::Axis(item) => {
                let id = item
                    .summary
                    .axes
                    .get(item.index as usize)
                    .map(|axis| axis.id)
                    .ok_or_else(|| error(Code::PathNotFound, "the axis is missing"))?;
                (item.chart.clone(), SetTarget::Axis(id))
            },
            _ => return Ok(None),
        };
        let Some(slide) = tree.container.slide() else {
            return Err(error(
                Code::UnsupportedCapability,
                "master and layout charts are not editable",
            ));
        };
        if tree.handle.parent().is_some() {
            return Err(error(
                Code::UnsupportedCapability,
                "chart writes require a top-level slide chart",
            ));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let (edit, geometry) = admit_set(self, &which, &tree, props, limits)?;
        let position = tree.handle.position().get();
        let source = self.editable_owner_for_write()?;
        let base = eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        if geometry.iter().any(Option::is_some) {
            let current = tree
                .handle
                .view()
                .common()
                .bounds()
                .map(|bounds| (bounds.x(), bounds.y(), bounds.width(), bounds.height()))
                .ok_or_else(|| {
                    error(
                        Code::UnsupportedCapability,
                        "the chart frame has no transform",
                    )
                })?;
            let bounds = (
                geometry[0].unwrap_or(current.0),
                geometry[1].unwrap_or(current.1),
                geometry[2].unwrap_or(current.2),
                geometry[3].unwrap_or(current.3),
            );
            native.set_shape_properties(
                slide as usize,
                &[
                    litchi_pptx::opened::ShapePropertiesReplacement::new(Position::new(position))
                        .with_bounds(bounds),
                ],
            )?;
        }
        let (xml, workbook) = native.chart(slide as usize, position)?;
        if let Some(edited) = edit.apply(xml, true)? {
            // The embedded workbook follows a data or series-name edit when
            // the chart refers to the authored layout; a data edit of any
            // other embedded workbook is refused, so the workbook never
            // silently diverges from the chart's cached values.
            let rebuilt = edited.spec.clone().map(|spec| {
                spec.with_binding(litchi_drawingml::chart::authoring::Binding::sheet("Sheet1"))
            });
            let workbook = match workbook {
                Some(_) if edited.data_changed || edited.names_changed => {
                    match rebuilt.or_else(|| shared::embedded_spec(&edited.xml)) {
                        Some(spec) => Some(
                            litchi_pptx::presentation::embedded::excel::Workbook::from_spec(&spec)?
                                .into_bytes(),
                        ),
                        None if edited.data_changed => {
                            return Err(error(
                                Code::UnsupportedCapability,
                                "the chart's embedded workbook is not in the authored layout, so its data cannot follow the edit",
                            )
                            .with_suggestion(
                                "edit the chart's data in PowerPoint, or set only titles, names and the legend",
                            ));
                        },
                        None => None,
                    }
                },
                _ => None,
            };
            native.set_chart(slide as usize, position, edited.xml, workbook)?;
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        // An exact no-op keeps this handle and its owner.
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Some(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        )))
    }
}

/// The eager package of an editable owner, promoting a source owner.
fn eager(source: &Editable, limits: litchi_opc::SourceLimits) -> Result<Arc<litchi_pptx::Package>> {
    Ok(match source {
        Editable::Eager(package) => Arc::clone(package),
        Editable::Source(package) => super::eager_package_from_source(package.to_bytes()?, limits)?,
    })
}

/// Admit a chart, series or axis `set`: the chart edit, and the frame
/// geometry `[x, y, width, height]` it moves to.
fn admit_set(
    file: &Pptx,
    which: &SetTarget,
    tree: &TreeItem,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<(ChartEdit, [Option<i64>; 4])> {
    let mut edit = ChartEdit::default();
    let mut geometry = [None; 4];
    let mut staged_bytes = 0_usize;
    let mut charge = |value: &Option<Scalar<'_>>| -> Result<()> {
        if let Some(Scalar::Text(value)) = value {
            staged_bytes = staged_bytes.saturating_add(value.len());
            bound(
                LimitKind::StagedBytes,
                staged_bytes,
                limits.max_staged_bytes(),
            )?;
        }
        Ok(())
    };
    let mut summary: Option<Arc<Summary>> = None;
    let chart = |summary: &mut Option<Arc<Summary>>| -> Result<Arc<Summary>> {
        if let Some(summary) = summary {
            return Ok(Arc::clone(summary));
        }
        let read = chart_summary(file, tree)?;
        *summary = Some(Arc::clone(&read));
        Ok(read)
    };
    for (key, value) in props.iter() {
        match which {
            SetTarget::Chart => {
                let (index, entry) = CHART_WRITE_TABLE
                    .resolve(key.as_str())
                    .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                let name = entry.spec.name;
                entry.prepare(Operation::Set, value, |setter, scalar| {
                    charge(&scalar)?;
                    match setter {
                        ChartSetter::Title => edit.title = Some(text(scalar)),
                        ChartSetter::Legend => {
                            edit.legend =
                                Some(shared::parse_legend(&text(scalar).unwrap_or_default())?);
                        },
                        ChartSetter::Categories => {
                            edit.categories = Some(shared::parse_list(
                                &text(scalar).unwrap_or_default(),
                                "categories",
                            )?);
                        },
                        ChartSetter::CatTitle | ChartSetter::AxisTitle => {
                            let role = if matches!(setter, ChartSetter::CatTitle) {
                                Role::Category
                            } else {
                                Role::Value
                            };
                            let summary = chart(&mut summary)?;
                            let axis = summary.axis(role).ok_or_else(|| {
                                error(Code::UnsupportedCapability, "the chart has no such axis")
                                    .with_property(name)
                            })?;
                            edit.axis_titles.push((axis.id, text(scalar)));
                        },
                        ChartSetter::Dotted(number, field) => {
                            let summary = chart(&mut summary)?;
                            edit.dotted_series(
                                &summary,
                                usize::from(number) + 1,
                                field,
                                text(scalar),
                            )?;
                        },
                        ChartSetter::ChartType => {
                            let summary = chart(&mut summary)?;
                            edit.chart_type(&summary, &text(scalar).unwrap_or_default())?;
                        },
                        ChartSetter::Data => {
                            let summary = chart(&mut summary)?;
                            edit.data(&summary, &text(scalar).unwrap_or_default())?;
                        },
                        ChartSetter::Format(key) => {
                            let summary = chart(&mut summary)?;
                            edit.format(&summary, key, name, scalar)?;
                        },
                        ChartSetter::X
                        | ChartSetter::Y
                        | ChartSetter::Width
                        | ChartSetter::Height => {
                            let Some(Scalar::Length(length)) = scalar else {
                                unreachable!("chart geometry coercion")
                            };
                            let slot = match setter {
                                ChartSetter::X => 0,
                                ChartSetter::Y => 1,
                                ChartSetter::Width => 2,
                                _ => 3,
                            };
                            geometry[slot] = Some(length.emu());
                        },
                        ChartSetter::Name | ChartSetter::Numbered(_) => {
                            unreachable!("chart set admission")
                        },
                    }
                    Ok(())
                })?;
                let _ = index;
            },
            SetTarget::Series(index) => {
                let (_, entry) = SERIES_WRITE_TABLE
                    .resolve(key.as_str())
                    .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                entry.prepare(Operation::Set, value, |setter, scalar| {
                    charge(&scalar)?;
                    match setter {
                        SeriesSetter::Name => edit.series_names.push((*index, text(scalar))),
                        SeriesSetter::Values => edit.series_values.push((
                            *index,
                            shared::parse_values(&text(scalar).unwrap_or_default(), "values")?,
                        )),
                        SeriesSetter::Color => edit.series_colors.push((
                            *index,
                            text(scalar)
                                .map(|color| shared::parse_color(&color, "color"))
                                .transpose()?,
                        )),
                    }
                    Ok(())
                })?;
            },
            SetTarget::Axis(id) => {
                let (_, entry) = AXIS_WRITE_TABLE
                    .resolve(key.as_str())
                    .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                entry.prepare(Operation::Set, value, |setter, scalar| {
                    charge(&scalar)?;
                    match setter {
                        AxisSetter::Title => edit.axis_titles.push((*id, text(scalar))),
                        AxisSetter::Min | AxisSetter::Max => {
                            let summary = chart(&mut summary)?;
                            let axis =
                                summary.axes.iter().find(|axis| axis.id == *id).ok_or_else(
                                    || error(Code::PathNotFound, "the axis is missing"),
                                )?;
                            edit.axis_bound(
                                axis,
                                matches!(setter, AxisSetter::Max),
                                number(scalar),
                            )?;
                        },
                    }
                    Ok(())
                })?;
            },
        }
    }
    Ok((edit, geometry))
}

fn number(value: Option<Scalar<'_>>) -> Option<f64> {
    match value {
        Some(Scalar::Number(value)) => Some(value),
        Some(Scalar::Int(value)) => Some(value as f64),
        _ => None,
    }
}

fn chart_summary(file: &Pptx, tree: &TreeItem) -> Result<Arc<Summary>> {
    tree.chart(file)?
        .ok_or_else(|| error(Code::CorruptFile, "the chart frame names no chart part"))
}

/// The children a chart lists: its series, then its axes.
pub(super) fn children(file: &Pptx, tree: &TreeItem) -> Result<Vec<Item>> {
    let Some(summary) = tree.chart(file)? else {
        return Ok(Vec::new());
    };
    let series = (0..summary.series.len()).map(|index| {
        Item::Series(ChartItem {
            chart: tree.clone(),
            summary: Arc::clone(&summary),
            index: u32::try_from(index).unwrap_or(u32::MAX),
        })
    });
    let axes = (0..summary.axes.len()).map(|index| {
        Item::Axis(ChartItem {
            chart: tree.clone(),
            summary: Arc::clone(&summary),
            index: u32::try_from(index).unwrap_or(u32::MAX),
        })
    });
    Ok(series.chain(axes).collect())
}
