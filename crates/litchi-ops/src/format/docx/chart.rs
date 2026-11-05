//! Typed chart CRUD in Word documents (record 0283).
//!
//! - `add(/body/p[N], chart, {chartType, title, data, categories, legend,
//!   catTitle, axisTitle, name, alt, width, height, ...})` builds an
//!   authored chart spec ([`crate::format::chart`], the same vocabulary as
//!   PPTX and XLSX) and inserts it as an inline drawing through the leaf
//!   `litchi_docx::document::Edit::insert_chart_at`: the chart part (or an
//!   extended `cx:chartSpace` part with its companions), a deterministic
//!   embedded workbook and the run, in one commit.
//! - Charts are addressed in their paragraph (`/body/p[N]/chart[M]`, the
//!   canonical path) or document-wide in document order, as OfficeCLI
//!   returns them (`/chart[N]`, also `/body/chart[N]`).
//! - `get` reads the frame (name, alternative text, extent) and the chart
//!   part (type, title, legend, categories, series, formatting keys).
//! - `set` edits titles, the legend, categories, series and axes through
//!   the source-preserving chart editor, rebuilds an authored chart for a
//!   new type, new data or a formatting key, resizes the frame, and keeps
//!   the embedded workbook in step with the chart's data; `remove` deletes
//!   the drawing (and its run when it holds nothing else).

use std::borrow::Cow;
use std::sync::{Arc, OnceLock};

use litchi_core::selector::Path;
use litchi_core::{FileFormat, Position};
use litchi_docx::chart::InlineChart;
use litchi_docx::document::NewChart;

use super::element::{element_path, path_with, refusal};
use super::node::{Address, Node, Step, Value};
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::format::chart::{self as shared, AddProps, ChartEdit, Role, Summary};
use crate::new_node::NewNode;
use crate::node::Value as NodeValue;
use crate::node::spec::PropSpec;
use crate::path::navigate::KindSet;
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::value::{LengthSpec, Scalar, ValueType};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

/// Chart assignments.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum ChartSetter {
    ChartType,
    Title,
    Legend,
    Categories,
    CatTitle,
    Name,
    Alt,
    Width,
    Height,
    Data,
    AxisTitle,
    /// A formatting key ([`shared::FORMAT_PROPS`]).
    Format(shared::FormatKey),
    /// `series{N}`, by its zero-based number.
    Numbered(u8),
    /// `series{N}.name`, `.values` or `.color`, by the zero-based number.
    Dotted(u8, shared::SeriesField),
}

const EXTENT: Constraint = Constraint::Length {
    min: crate::value::Length::from_emu(1),
    max: crate::value::Length::from_emu(22 * 914_400),
};

const fn entry<S: Copy>(
    spec: PropSpec,
    access: Access<S>,
    constraint: Constraint,
    clear: bool,
) -> Entry<S> {
    Entry {
        spec,
        access,
        constraint,
        clear,
    }
}

const CHART_BASE_COUNT: usize = 12;
const CHART_FORMAT_START: usize = CHART_BASE_COUNT;
const CHART_NUMBERED_START: usize = CHART_FORMAT_START + shared::FORMAT_PROPS.len();
const CHART_DOTTED_START: usize = CHART_NUMBERED_START + shared::NUMBERED_SERIES_PROPS.len();
const CHART_WRITE_COUNT: usize = CHART_DOTTED_START + shared::DOTTED_SERIES_PROPS.len();
const _: () = assert!(
    CHART_WRITE_COUNT <= shared::MAX_TABLE,
    "chart adds track assignments in an Assigned set"
);

const fn chart_entries() -> [Entry<ChartSetter>; CHART_WRITE_COUNT] {
    let base: [Entry<ChartSetter>; CHART_BASE_COUNT] = [
        entry(
            PropSpec::new("chartType", ValueType::Text),
            Access::SetAndAdd(ChartSetter::ChartType),
            Constraint::None,
            false,
        ),
        entry(
            PropSpec::new("title", ValueType::Text),
            Access::SetAndAdd(ChartSetter::Title),
            Constraint::None,
            true,
        ),
        entry(
            PropSpec::new("legend", ValueType::Text),
            Access::SetAndAdd(ChartSetter::Legend),
            Constraint::None,
            false,
        ),
        entry(
            PropSpec::new("categories", ValueType::Text),
            Access::SetAndAdd(ChartSetter::Categories),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            PropSpec::new("catTitle", ValueType::Text),
            Access::SetAndAdd(ChartSetter::CatTitle),
            Constraint::None,
            true,
        ),
        Entry::read(PropSpec::new("seriesCount", ValueType::Int)),
        entry(
            PropSpec::new("name", ValueType::Text),
            Access::Add(ChartSetter::Name),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            PropSpec::new("alt", ValueType::Text),
            Access::Add(ChartSetter::Alt),
            Constraint::None,
            false,
        ),
        entry(
            PropSpec::new("width", ValueType::Length(LengthSpec::EXTENT)),
            Access::SetAndAdd(ChartSetter::Width),
            EXTENT,
            false,
        ),
        entry(
            PropSpec::new("height", ValueType::Length(LengthSpec::EXTENT)),
            Access::SetAndAdd(ChartSetter::Height),
            EXTENT,
            false,
        ),
        entry(
            write_only("data", ValueType::Text),
            Access::SetAndAdd(ChartSetter::Data),
            Constraint::NotEmpty,
            false,
        ),
        entry(
            write_only("axisTitle", ValueType::Text),
            Access::SetAndAdd(ChartSetter::AxisTitle),
            Constraint::None,
            true,
        ),
    ];
    let mut entries = [base[0]; CHART_WRITE_COUNT];
    let mut index = 0;
    while index < CHART_BASE_COUNT {
        entries[index] = base[index];
        index += 1;
    }
    let mut format = 0;
    while format < shared::FORMAT_PROPS.len() {
        let (key, spec) = shared::FORMAT_PROPS[format];
        entries[CHART_FORMAT_START + format] = entry(
            spec,
            Access::SetAndAdd(ChartSetter::Format(key)),
            Constraint::None,
            false,
        );
        format += 1;
    }
    let mut number = 0;
    while number < shared::NUMBERED_SERIES_PROPS.len() {
        entries[CHART_NUMBERED_START + number] = entry(
            shared::NUMBERED_SERIES_PROPS[number],
            Access::Add(ChartSetter::Numbered(number as u8)),
            Constraint::NotEmpty,
            false,
        );
        number += 1;
    }
    let mut dotted = 0;
    while dotted < shared::DOTTED_SERIES_PROPS.len() {
        let field = shared::SERIES_FIELDS[dotted % 3];
        entries[CHART_DOTTED_START + dotted] = entry(
            shared::DOTTED_SERIES_PROPS[dotted],
            Access::SetAndAdd(ChartSetter::Dotted((dotted / 3) as u8, field)),
            Constraint::NotEmpty,
            false,
        );
        dotted += 1;
    }
    entries
}

pub(crate) const CHART_WRITE_TABLE: Table<ChartSetter, CHART_WRITE_COUNT, 4> = Table::new(
    chart_entries(),
    [
        Alias {
            name: "type",
            entry: 0,
        },
        Alias {
            name: "htitle",
            entry: 4,
        },
        Alias {
            name: "vtitle",
            entry: 11,
        },
        Alias {
            name: "description",
            entry: 7,
        },
    ],
);

/// The chart element's properties: the write table, readable where `get`
/// is set.
pub(crate) const CHART_PROPS: &[PropSpec] = &CHART_WRITE_TABLE.readable();

#[derive(Clone, Copy)]
pub(crate) enum SeriesSetter {
    Name,
    Values,
    Color,
}

/// A chart series (`/body/p[N]/chart[M]/series[K]`, record 0335): the
/// PPTX and XLSX series vocabulary.
pub(crate) const SERIES_WRITE_TABLE: Table<SeriesSetter, 4, 0> = Table::new(
    [
        entry(
            PropSpec::new("name", ValueType::Text),
            Access::Set(SeriesSetter::Name),
            Constraint::None,
            false,
        ),
        entry(
            PropSpec::new("values", ValueType::Text),
            Access::Set(SeriesSetter::Values),
            Constraint::NotEmpty,
            false,
        ),
        Entry::read(PropSpec::new("categories", ValueType::Text)),
        entry(
            PropSpec::new("color", ValueType::Text),
            Access::Set(SeriesSetter::Color),
            Constraint::NotEmpty,
            true,
        ),
    ],
    [],
);

/// The series element's properties.
pub(crate) const SERIES_PROPS: &[PropSpec] = &SERIES_WRITE_TABLE.readable();

#[derive(Clone, Copy)]
pub(crate) enum AxisSetter {
    Title,
    Min,
    Max,
}

/// A chart axis (`/body/p[N]/chart[M]/axis[@role=ROLE]`, record 0335).
pub(crate) const AXIS_WRITE_TABLE: Table<AxisSetter, 4, 0> = Table::new(
    [
        Entry::read(PropSpec::new("role", ValueType::Text)),
        entry(
            PropSpec::new("title", ValueType::Text),
            Access::Set(AxisSetter::Title),
            Constraint::None,
            true,
        ),
        entry(
            PropSpec::new("min", ValueType::Number),
            Access::Set(AxisSetter::Min),
            Constraint::None,
            true,
        ),
        entry(
            PropSpec::new("max", ValueType::Number),
            Access::Set(AxisSetter::Max),
            Constraint::None,
            true,
        ),
    ],
    [],
);

/// The axis element's properties.
pub(crate) const AXIS_PROPS: &[PropSpec] = &AXIS_WRITE_TABLE.readable();

/// One chart of a paragraph, as a node carries it.
pub(crate) struct ChartValue {
    /// The paragraph's own node value and address, for its element path.
    pub(super) paragraph: Value,
    pub(super) paragraph_address: Address,
    /// The chart's position among the paragraph's charts.
    pub(super) ordinal: u32,
    pub(super) chart: InlineChart,
    /// The chart part's projection, read on first use.
    summary: OnceLock<Option<Arc<Summary>>>,
}

/// The node of the chart at `ordinal` of a paragraph.
pub(super) fn chart_node<'f>(
    docx: &'f Docx,
    paragraph: &Value,
    paragraph_address: &Address,
    ordinal: u32,
    chart: InlineChart,
) -> Node<'f> {
    Node::new(
        docx,
        Kind::Chart,
        Value::Chart(Arc::new(ChartValue {
            paragraph: paragraph.clone(),
            paragraph_address: paragraph_address.clone(),
            ordinal,
            chart,
            summary: OnceLock::new(),
        })),
        paragraph_address.with(Step::new(Kind::Chart, ordinal)),
    )
}

impl ChartValue {
    /// The paragraph node holding the chart.
    pub(super) fn paragraph_node<'f>(&self, docx: &'f Docx) -> Node<'f> {
        Node::new(
            docx,
            Kind::Paragraph,
            self.paragraph.clone(),
            self.paragraph_address.clone(),
        )
    }

    /// The chart part's projection, read once per handle.
    pub(super) fn summary(&self, docx: &Docx) -> Result<Option<Arc<Summary>>> {
        if let Some(summary) = self.summary.get() {
            return Ok(summary.clone());
        }
        let id = self.chart.relationship_id();
        let cached = docx.parts().charts.lock().ok().and_then(|charts| {
            charts
                .iter()
                .find(|(cached, _)| cached.as_ref() == id)
                .map(|(_, summary)| summary.clone())
        });
        let summary = match cached {
            Some(summary) => summary,
            None => {
                let result = match docx.owner() {
                    DocxOwner::Eager(package) => package.chart_xml(id).map_err(Error::from),
                    DocxOwner::Source(package) => package.chart_xml(id).map_err(Error::from),
                };
                let xml = docx.owner().finish(result)?;
                let summary = xml
                    .map(|xml| Summary::read(&xml).map(Arc::new))
                    .transpose()?;
                if let Ok(mut charts) = docx.parts().charts.lock() {
                    charts.push((id.into(), summary.clone()));
                }
                summary
            },
        };
        let _ = self.summary.set(summary.clone());
        Ok(summary)
    }
}

/// The charts of the main document in document order: every paragraph of
/// the body, its tables' cells and its content controls, walked depth-first.
pub(super) fn document_charts(docx: &Docx) -> Result<Vec<Node<'_>>> {
    let containers = KindSet::of(Kind::Paragraph)
        .with(Kind::Table)
        .with(Kind::Row)
        .with(Kind::Cell)
        .with(Kind::Sdt);
    let mut found = Vec::new();
    let body = super::nav::body_node(docx)?;
    let mut stack = vec![body];
    // Depth-first, children in order: push them reversed.
    while let Some(node) = stack.pop() {
        if node.kind == Kind::Paragraph {
            let charts = super::children::Children::of(&node, KindSet::of(Kind::Chart))?;
            for chart in charts {
                found.push(chart?);
            }
            continue;
        }
        let mut children = Vec::new();
        for child in super::children::Children::of(&node, containers)? {
            children.push(child?);
        }
        stack.extend(children.into_iter().rev());
    }
    Ok(found)
}

/// The document-wide chart a `/chart[N]` (or `/body/chart[N]`) step picks.
pub(super) fn pick_document_chart<'f>(
    docx: &'f Docx,
    select: &crate::path::navigate::Select<'_>,
) -> Result<crate::path::navigate::Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    use crate::path::navigate::{Lookup, Select};
    let charts = document_charts(docx)?;
    let picked = match *select {
        Select::None => Some(0),
        Select::Ordinal(position) => Some(position.get()),
        Select::Last => charts.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(
        match picked.and_then(|index| charts.into_iter().nth(index)) {
            Some(node) => Lookup::Found(node),
            None => Lookup::Missing,
        },
    )
}

/// The series and axes of a chart node, in the order PPTX and XLSX list
/// them: the series, then the axes (record 0335).
pub(super) fn chart_children(
    docx: &Docx,
    chart: &Arc<ChartValue>,
    kinds: KindSet<Kind>,
) -> Result<Vec<(Kind, Value, Step)>> {
    let Some(summary) = chart.summary(docx)? else {
        return Ok(Vec::new());
    };
    let mut out = Vec::new();
    for (kind, count) in [
        (Kind::Series, summary.series.len()),
        (Kind::Axis, summary.axes.len()),
    ] {
        if !kinds.contains(kind) {
            continue;
        }
        for index in 0..count {
            let ordinal = u32::try_from(index).unwrap_or(u32::MAX);
            out.push((
                kind,
                Value::ChartChild(Arc::clone(chart), ordinal),
                Step::new(kind, ordinal),
            ));
        }
    }
    Ok(out)
}

/// The series (by position) or axis (by position or `@role`) a step picks
/// under a chart.
pub(super) fn pick_chart_child<'f>(
    docx: &'f Docx,
    parent: &Node<'f>,
    chart: &Arc<ChartValue>,
    kind: Kind,
    select: &crate::path::navigate::Select<'_>,
) -> Result<crate::path::navigate::Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    use crate::path::navigate::{Lookup, Select};
    let Some(summary) = chart.summary(docx)? else {
        return Ok(Lookup::Missing);
    };
    let count = if kind == Kind::Series {
        summary.series.len()
    } else {
        summary.axes.len()
    };
    let index = match select {
        Select::Ordinal(ordinal) => Some(ordinal.get()),
        Select::None => Some(0),
        Select::Last => count.checked_sub(1),
        Select::Identity(identities) if kind == Kind::Axis => {
            let mut found = None;
            for (identity, value) in identities.iter() {
                if !identity.key.eq_ignore_ascii_case("role") {
                    return Ok(Lookup::NotApplicable);
                }
                found = Role::parse(value)
                    .and_then(|role| summary.axes.iter().position(|axis| axis.role == role));
            }
            found
        },
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match index.filter(|index| *index < count) {
        Some(index) => {
            let ordinal = u32::try_from(index).unwrap_or(u32::MAX);
            Lookup::Found(Node::new(
                docx,
                kind,
                Value::ChartChild(Arc::clone(chart), ordinal),
                parent.child_address().with(Step::new(kind, ordinal)),
            ))
        },
        None => Lookup::Missing,
    })
}

/// The value of a series node's entry at `index` (`name`, `values`,
/// `categories`, `color`).
pub(super) fn series_value(
    docx: &Docx,
    chart: &ChartValue,
    child: u32,
    index: usize,
) -> Option<NodeValue<'static>> {
    let summary = chart.summary(docx).ok()??;
    let series = summary.series.get(child as usize)?;
    let text = |text: &str| NodeValue::Text(Cow::Owned(text.to_owned()));
    match index {
        0 => series.name.as_deref().map(text),
        1 => Some(text(&shared::join_values(&series.values))),
        2 => (!series.categories.is_empty()).then(|| text(&series.categories.join(","))),
        3 => series.color.as_deref().map(text),
        _ => None,
    }
}

/// The value of an axis node's entry at `index` (`role`, `title`, `min`,
/// `max`).
pub(super) fn axis_value(
    docx: &Docx,
    chart: &ChartValue,
    child: u32,
    index: usize,
) -> Option<NodeValue<'static>> {
    let summary = chart.summary(docx).ok()??;
    let axis = summary.axes.get(child as usize)?;
    match index {
        0 => Some(NodeValue::Text(Cow::Borrowed(axis.role.name()))),
        1 => axis
            .title
            .as_deref()
            .map(|title| NodeValue::Text(Cow::Owned(title.to_owned()))),
        2 => axis.min.and_then(NodeValue::number),
        3 => axis.max.and_then(NodeValue::number),
        _ => None,
    }
}

/// The number of a chart's series and axes.
pub(super) fn child_count(docx: &Docx, chart: &ChartValue) -> Result<usize> {
    Ok(chart
        .summary(docx)?
        .map_or(0, |summary| summary.series.len() + summary.axes.len()))
}

fn length(emu: i64) -> NodeValue<'static> {
    NodeValue::Length(crate::value::Length::from_emu(emu))
}

/// The value of the chart table's entry at `index`.
pub(super) fn chart_value(
    docx: &Docx,
    chart: &ChartValue,
    index: usize,
) -> Option<NodeValue<'static>> {
    let text = |text: &str| NodeValue::Text(Cow::Owned(text.to_owned()));
    match index {
        6 => (!chart.chart.name().is_empty()).then(|| text(chart.chart.name())),
        7 => (!chart.chart.description().is_empty()).then(|| text(chart.chart.description())),
        8 => Some(length(chart.chart.width_emu())),
        9 => Some(length(chart.chart.height_emu())),
        _ => {
            let summary = chart.summary(docx).ok()??;
            match index {
                0 => summary.chart_type().map(|name| text(&name)),
                1 => summary.title.as_deref().map(text),
                2 => Some(NodeValue::Text(Cow::Borrowed(shared::legend_name(
                    summary.legend,
                )))),
                3 => summary
                    .categories()
                    .map(|categories| text(&categories.join(","))),
                4 => summary
                    .axis(Role::Category)
                    .and_then(|axis| axis.title.as_deref())
                    .map(text),
                5 => Some(NodeValue::Int(
                    i64::try_from(summary.series.len()).unwrap_or(i64::MAX),
                )),
                _ => {
                    let format = index.checked_sub(CHART_FORMAT_START)?;
                    let (key, _) = shared::FORMAT_PROPS.get(format)?;
                    summary.format(*key)
                },
            }
        },
    }
}

fn text(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(value) | Scalar::Token(value)) => Some(value.to_owned()),
        _ => None,
    }
}

fn number(value: Option<Scalar<'_>>) -> Option<f64> {
    match value {
        Some(Scalar::Number(number)) => Some(number),
        Some(Scalar::Int(number)) => Some(number as f64),
        _ => None,
    }
}

fn emu(value: Option<Scalar<'_>>) -> Option<i64> {
    match value {
        Some(Scalar::Length(length)) => Some(length.emu()),
        _ => None,
    }
}

impl Docx {
    /// Add a chart to a paragraph: `/body/p[N]` (or a paragraph in a cell or
    /// a content control).
    pub(super) fn add_chart(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
        direct: Option<Position>,
    ) -> Result<(Commit, Path<'static>)> {
        let paragraph = element_path(parent)?;
        let mut props = AddProps::default();
        let mut name = None;
        let mut alt = None;
        let mut size = (None, None);
        let mut assigned = shared::Assigned::default();
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
                match setter {
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
                    ChartSetter::CatTitle => props.category_axis_title = text(scalar),
                    ChartSetter::AxisTitle => props.value_axis_title = text(scalar),
                    ChartSetter::Data => {
                        props.data = Some(shared::parse_data(&text(scalar).unwrap_or_default())?);
                    },
                    ChartSetter::Name => name = text(scalar),
                    ChartSetter::Alt => alt = text(scalar),
                    ChartSetter::Width => size.0 = emu(scalar),
                    ChartSetter::Height => size.1 = emu(scalar),
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
        for value in [&name, &alt].into_iter().flatten() {
            if value.contains('\0') {
                return Err(error(Code::InvalidValue, "chart text cannot contain NUL"));
            }
        }
        let spec = props.into_spec()?;
        // OfficeCLI names an unnamed chart after its title.
        let name = name.or_else(|| spec.title().map(str::to_owned));
        let mut chart = NewChart::new(spec).with_size(size.0, size.1);
        if let Some(name) = name {
            chart = chart.with_name(name);
        }
        if let Some(alt) = alt {
            chart = chart.with_description(alt);
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("charts need the owning package editor"));
        };
        let ordinal = eager
            .insert_chart_at(&paragraph, direct, &chart)
            .map_err(Error::from)?;
        let commit = edit.commit(self)?;
        Ok((commit, path_with(parent, Some((Kind::Chart, ordinal)))))
    }

    /// A `set` whose targets are charts; `None` for other targets.
    pub(super) fn try_set_chart(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty()
            || nodes
                .iter()
                .any(|node| !matches!(node.kind, Kind::Chart | Kind::Series | Kind::Axis))
        {
            return Ok(Err(editor));
        }
        if nodes.len() != 1 {
            return Err(error(
                Code::UnsupportedCapability,
                "a chart set targets one chart, series or axis",
            ));
        }
        if let Value::ChartChild(chart, child) = &nodes[0].value {
            let edit = self.chart_child_edit(chart, nodes[0].kind, *child, props)?;
            return self
                .commit_chart_edit(editor, chart, edit, (None, None))
                .map(Ok);
        }
        let Value::Chart(chart) = &nodes[0].value else {
            return Err(error(Code::InternalError, "the chart node holds no chart"));
        };
        // The chart part is read only when a key needs it (a title or a
        // size does not).
        let summary = || -> Result<Arc<Summary>> {
            chart
                .summary(self)?
                .ok_or_else(|| error(Code::CorruptFile, "the chart names no chart part"))
        };
        let mut edit = ChartEdit::default();
        let mut size = (None, None);
        for (key, value) in props.iter() {
            let (_, entry) = CHART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let name = entry.spec.name;
            entry.prepare(Operation::Set, value, |setter, scalar| {
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
                        let summary = summary()?;
                        let axis = summary.axis(role).ok_or_else(|| {
                            error(Code::UnsupportedCapability, "the chart has no such axis")
                                .with_property(name)
                        })?;
                        edit.axis_titles.push((axis.id, text(scalar)));
                    },
                    ChartSetter::Dotted(number, field) => edit.dotted_series(
                        &*summary()?,
                        usize::from(number) + 1,
                        field,
                        text(scalar),
                    )?,
                    ChartSetter::ChartType => {
                        edit.chart_type(&*summary()?, &text(scalar).unwrap_or_default())?;
                    },
                    ChartSetter::Data => {
                        edit.data(&*summary()?, &text(scalar).unwrap_or_default())?
                    },
                    ChartSetter::Format(key) => edit.format(&*summary()?, key, name, scalar)?,
                    ChartSetter::Width => size.0 = emu(scalar),
                    ChartSetter::Height => size.1 = emu(scalar),
                    ChartSetter::Name | ChartSetter::Alt | ChartSetter::Numbered(_) => {
                        unreachable!("chart set admission")
                    },
                }
                Ok(())
            })?;
        }
        self.commit_chart_edit(editor, chart, edit, size).map(Ok)
    }

    /// The edit of a `set` on a chart's series or axis (record 0335).
    fn chart_child_edit(
        &self,
        chart: &ChartValue,
        kind: Kind,
        child: u32,
        props: &Props<'_>,
    ) -> Result<ChartEdit> {
        let summary = chart
            .summary(self)?
            .ok_or_else(|| error(Code::CorruptFile, "the chart names no chart part"))?;
        let mut edit = ChartEdit::default();
        for (key, value) in props.iter() {
            if kind == Kind::Series {
                let index = summary
                    .series
                    .get(child as usize)
                    .map(|series| series.index)
                    .ok_or_else(|| error(Code::PathNotFound, "the series is missing"))?;
                let (_, entry) = SERIES_WRITE_TABLE
                    .resolve(key.as_str())
                    .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                entry.prepare(Operation::Set, value, |setter, scalar| {
                    match setter {
                        SeriesSetter::Name => edit.series_names.push((index, text(scalar))),
                        SeriesSetter::Values => edit.series_values.push((
                            index,
                            shared::parse_values(&text(scalar).unwrap_or_default(), "values")?,
                        )),
                        SeriesSetter::Color => edit.series_colors.push((
                            index,
                            text(scalar)
                                .map(|color| shared::parse_color(&color, "color"))
                                .transpose()?,
                        )),
                    }
                    Ok(())
                })?;
            } else {
                let axis = summary
                    .axes
                    .get(child as usize)
                    .ok_or_else(|| error(Code::PathNotFound, "the axis is missing"))?;
                let (_, entry) = AXIS_WRITE_TABLE
                    .resolve(key.as_str())
                    .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                entry.prepare(Operation::Set, value, |setter, scalar| {
                    match setter {
                        AxisSetter::Title => edit.axis_titles.push((axis.id, text(scalar))),
                        AxisSetter::Min | AxisSetter::Max => edit.axis_bound(
                            axis,
                            matches!(setter, AxisSetter::Max),
                            number(scalar),
                        )?,
                    }
                    Ok(())
                })?;
            }
        }
        Ok(edit)
    }

    /// Stage a chart edit and a frame resize, and commit them.
    fn commit_chart_edit(
        &self,
        editor: DocxEditor,
        chart: &ChartValue,
        edit: ChartEdit,
        size: (Option<i64>, Option<i64>),
    ) -> Result<Commit> {
        let paragraph = element_path(&chart.paragraph_node(self))?;
        let ordinal = Position::new(chart.ordinal as usize);
        self.owner().check_source()?;
        let mut work = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut work else {
            return Err(refusal("charts need the owning package editor"));
        };
        if size.0.is_some() || size.1.is_some() {
            let extent = (
                size.0.unwrap_or(chart.chart.width_emu()),
                size.1.unwrap_or(chart.chart.height_emu()),
            );
            eager
                .set_chart_extent_at(&paragraph, ordinal, extent)
                .map_err(Error::from)?;
        }
        let id = chart.chart.relationship_id();
        let (xml, workbook) = eager
            .chart_parts(id)
            .map_err(Error::from)?
            .ok_or_else(|| error(Code::CorruptFile, "the chart names no chart part"))?;
        if let Some(edited) = edit.apply(xml, true)? {
            let workbook = match workbook {
                Some(_) if edited.data_changed || edited.names_changed => {
                    let spec = edited
                        .spec
                        .clone()
                        .map(|spec| {
                            spec.with_binding(litchi_drawingml::chart::authoring::Binding::sheet(
                                "Sheet1",
                            ))
                        })
                        .or_else(|| shared::embedded_spec(&edited.xml));
                    match spec {
                        Some(spec) => Some(
                            litchi_docx::document::chart_workbook(&spec).map_err(Error::from)?,
                        ),
                        None if edited.data_changed => {
                            return Err(error(
                                Code::UnsupportedCapability,
                                "the chart's embedded workbook is not in the authored layout, so its data cannot follow the edit",
                            )
                            .with_suggestion(
                                "edit the chart's data in Word, or set only titles, names and the legend",
                            ));
                        },
                        None => None,
                    }
                },
                _ => None,
            };
            eager
                .set_chart(id, edited.xml, workbook)
                .map_err(Error::from)?;
        }
        work.commit(self)
    }

    /// Remove charts selected by their paths.
    pub(super) fn remove_charts(&self, editor: DocxEditor, nodes: &[Node<'_>]) -> Result<Commit> {
        // Later charts first, so earlier ordinals stay valid.
        let mut targets = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::Chart(chart) = &node.value else {
                return Err(error(Code::InternalError, "the chart node holds no chart"));
            };
            targets.push((element_path(&chart.paragraph_node(self))?, chart.ordinal));
        }
        targets.sort_by(|left, right| {
            super::nested::order_key(&right.0)
                .cmp(&super::nested::order_key(&left.0))
                .then(right.1.cmp(&left.1))
        });
        targets.dedup_by(|left, right| left.0 == right.0 && left.1 == right.1);
        self.owner().check_source()?;
        let mut work = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut work else {
            return Err(refusal("charts need the owning package editor"));
        };
        for (paragraph, ordinal) in &targets {
            eager
                .remove_chart_at(paragraph, Position::new(*ordinal as usize))
                .map_err(Error::from)?;
        }
        work.commit(self)
    }
}

/// Where a chart `add` under a paragraph goes: before the paragraph's direct
/// run at `placement`, or at its end.
pub(super) fn placement_run(placement: Option<&Placement<'_>>) -> Result<()> {
    match placement {
        None | Some(Placement::End) => Ok(()),
        Some(_) => Err(error(
            Code::UnsupportedCapability,
            "chart insertion currently supports end placement",
        )
        .with_suggestion("omit placement or use Placement::End")),
    }
}
