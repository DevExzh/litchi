use super::super::super::{
    Kind, Ref, Stream, axis, cache as chart_cache, codec, format, group, layout,
};
use super::super::cache::{Cache, Value};
use super::super::context::{Context, Count, GroupId, Order, Props, Rect};
use super::super::groups::{Family, Group};
use super::super::inventory::{Edit, Label, Legend, Origin, Raw};
use super::super::series::{Link, Owner, Role, RowCol, Series, Source};
use super::validation::{cache_dimensions, check_add, dimensions_cover, reserve_one};
use crate::{Error, Limits, Result};

/// Host-neutral semantic chart.
///
/// Parsed values retain their exact [`Stream`]. An untouched parsed value can
/// therefore be encoded byte-for-byte without copying. Mutation is allowed for
/// inspection workflows, but encoding such a value is refused until a future
/// lossless record editor can prove placement of every opaque record.
/// Fresh values can be assembled through the deliberately scoped standalone
/// Graph profile. Other fresh values return [`Error::UnsupportedAuthoring`],
/// so self-consistent but unsupported streams cannot escape the crate.
#[derive(Debug)]
pub struct Chart {
    pub(in crate::chart) context: Context,
    pub(in crate::chart) rect: Rect,
    pub(in crate::chart) props: Props,
    pub(in crate::chart) zoom: layout::Zoom,
    pub(in crate::chart) growth: layout::Growth,
    pub(in crate::chart) title: Option<String>,
    pub(in crate::chart) series: Vec<Series>,
    pub(in crate::chart) groups: Vec<Group>,
    pub(in crate::chart) axes: Vec<axis::Axis>,
    pub(in crate::chart) parents: Vec<axis::Parent>,
    pub(in crate::chart) legend: Option<Legend>,
    pub(in crate::chart) caches: Vec<Cache>,
    pub(in crate::chart) dimensions: chart_cache::Dims,
    pub(in crate::chart) formats: Vec<format::Format>,
    pub(in crate::chart) labels: Vec<Label>,
    pub(in crate::chart) unknown: Vec<Raw>,
    pub(in crate::chart) origin: Origin,
    pub(in crate::chart) dirty: bool,
    pub(in crate::chart) limits: Limits,
    /// Internal proof gate for the record encoder. The public Graph builder
    /// enables it only after its scoped profile has been validated.
    pub(in crate::chart) authoring_proven: bool,
    /// Whether the public standalone Graph profile has proved its canonical
    /// datasheet orientation and record scaffold.
    pub(in crate::chart) graph_authoring_profile: bool,
}

/// The deliberately small standalone Graph authoring profile.
///
/// The profile has one primary chart group. Bar and line charts use one
/// category and one value axis; pie charts use the same primary axis group but
/// contain no axes, as required by [MS-OGRAPH] for pie chart groups. Other
/// chart families, secondary axes, and combination charts remain outside this
/// constructor's proof boundary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum GraphFamily {
    /// A bounded Bar chart group.
    Bar,
    /// A bounded Line chart group.
    Line,
    /// A bounded Pie chart group with no axes.
    Pie,
}

impl Chart {
    /// Creates a fresh chart using conservative limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] if the limits are inconsistent or
    /// [`Error::Allocation`] if an initial collection cannot be allocated.
    pub fn new(context: Context) -> Result<Self> {
        Self::new_with(context, Limits::default())
    }

    /// Creates a fresh chart with explicit authoring bounds.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] if the limits are inconsistent or
    /// [`Error::Allocation`] if an initial collection cannot be allocated.
    pub fn new_with(context: Context, limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        let mut groups = Vec::new();
        groups.try_reserve_exact(1).ok().ok_or(Error::Allocation {
            resource: "chart groups",
        })?;
        groups.push(Group::line());
        let mut parents = Vec::new();
        parents.try_reserve_exact(1).ok().ok_or(Error::Allocation {
            resource: "axis parents",
        })?;
        parents.push(axis::Parent::primary(layout::Pos::default()));
        Ok(Self {
            context,
            rect: Rect::default(),
            props: Props::default(),
            zoom: layout::Zoom::default(),
            growth: layout::Growth::default(),
            title: None,
            series: Vec::new(),
            groups,
            axes: Vec::new(),
            parents,
            legend: None,
            caches: Vec::new(),
            dimensions: chart_cache::Dims::empty(context.kind()),
            formats: Vec::new(),
            labels: Vec::new(),
            unknown: Vec::new(),
            origin: Origin::Fresh,
            dirty: false,
            limits: validated,
            authoring_proven: false,
            graph_authoring_profile: false,
        })
    }

    /// Creates the complete scaffold used by the standalone Graph authoring
    /// profile.
    ///
    /// This only creates the checked structural baseline. Call
    /// [`Self::finish_graph_authoring`] after adding the series and caches; the
    /// encoder will not emit a fresh chart until that final validation has
    /// succeeded.
    pub fn new_graph_authoring(family: GraphFamily, limits: Limits) -> Result<Self> {
        let mut chart = Self::new_with(Context::graph(), limits)?;
        let group = chart.groups.first_mut().ok_or(Error::InvalidModel {
            field: "groups",
            reason: "fresh chart did not contain its primary chart group",
        })?;
        group.family = match family {
            GraphFamily::Bar => Family::Bar {
                overlap: group::Overlap::ZERO,
                gap: group::Gap::new(150).ok_or(Error::InvalidModel {
                    field: "group",
                    reason: "default bar gap is outside its checked range",
                })?,
                // fTranspose=1 is the horizontal Bar profile; it avoids the
                // separate column/Chart3d series-axis combination.
                flags: 1,
            },
            // A stacked line group is the smallest positive Line profile
            // without a SeriesAxis/Chart3d pair. MS-OGRAPH requires that
            // pair when fStacked is zero, so the two-axis profile uses the
            // normative fStacked=1 form (with f100 left clear).
            GraphFamily::Line => Family::Line { flags: 1 },
            GraphFamily::Pie => Family::Pie {
                rotation: 0,
                hole: 0,
                flags: 0,
            },
        };
        if !matches!(family, GraphFamily::Pie) {
            chart.add_axis(axis::Axis::new(axis::Kind::Category))?;
            chart.add_axis(axis::Axis::new(axis::Kind::Value))?;
        } else {
            // PlotArea belongs to the optional AXES sequence. A pie profile
            // has no axes, so retaining the default PlotArea flag would emit
            // a record outside the normative Pie AXISPARENT grammar.
            chart.props.plot_area = false;
        }
        Ok(chart)
    }

    /// Completes and proves the standalone Graph authoring profile.
    ///
    /// The proof is established only after producer identity, collection
    /// ownership, chart-family/axis combinations, series links, cache
    /// dimensions, and all bounded record fields have been validated. Parsed
    /// charts and unsupported families cannot be promoted through this API.
    pub fn finish_graph_authoring(mut self) -> Result<Self> {
        self.validate_graph_authoring()?;
        self.authoring_proven = true;
        self.graph_authoring_profile = true;
        Ok(self)
    }

    fn validate_graph_authoring(&self) -> Result<()> {
        if self.context.kind() != Kind::Graph {
            return Err(Error::UnsupportedAuthoring {
                reason: "standalone Graph authoring requires a Graph producer context",
            });
        }
        if !matches!(&self.origin, Origin::Fresh) {
            return Err(Error::UnsupportedAuthoring {
                reason: "parsed chart replacement remains on the source-bound edit path",
            });
        }
        if self.parents.len() != 1 || self.parents[0].id() != axis::ParentId::PRIMARY {
            return Err(Error::UnsupportedAuthoring {
                reason: "the fresh Graph profile has exactly one primary AxisParent",
            });
        }
        if self.groups.len() != 1 {
            return Err(Error::UnsupportedAuthoring {
                reason: "fresh Graph authoring supports one chart group only",
            });
        }
        let group = self.groups.first().ok_or(Error::InvalidModel {
            field: "groups",
            reason: "fresh Graph chart group is missing",
        })?;
        if group.parent != axis::ParentId::PRIMARY
            || group.order != Order::ZERO
            || !group.lines.is_empty()
            || !group.drop_bars.is_empty()
        {
            return Err(Error::UnsupportedAuthoring {
                reason: "Graph authoring does not support secondary or opaque group children",
            });
        }
        let family = match group.family {
            Family::Bar {
                overlap,
                gap,
                flags,
            } if overlap == group::Overlap::ZERO && gap.get() == 150 && flags == 1 => {
                GraphFamily::Bar
            },
            Family::Line { flags: 1 } => GraphFamily::Line,
            Family::Pie {
                rotation,
                hole,
                flags,
            } if rotation == 0 && hole == 0 && flags == 0 => GraphFamily::Pie,
            Family::Area { .. }
            | Family::Scatter { .. }
            | Family::Radar { .. }
            | Family::Surface { .. }
            | Family::Bar { .. }
            | Family::Line { .. }
            | Family::Pie { .. } => {
                return Err(Error::UnsupportedAuthoring {
                    reason: "the selected Graph family settings are outside the authoring profile",
                });
            },
        };
        match family {
            GraphFamily::Pie if !self.axes.is_empty() => {
                return Err(Error::UnsupportedAuthoring {
                    reason: "pie chart groups must not contain axes",
                });
            },
            GraphFamily::Bar | GraphFamily::Line => {
                if self.axes.len() != 2
                    || self.axes[0].kind != axis::Kind::Category
                    || self.axes[1].kind != axis::Kind::Value
                    || self
                        .axes
                        .iter()
                        .any(|value| value.parent != axis::ParentId::PRIMARY)
                {
                    return Err(Error::UnsupportedAuthoring {
                        reason: "bar and line charts require one primary category and value axis",
                    });
                }
            },
            GraphFamily::Pie => {},
        }
        if matches!(family, GraphFamily::Bar | GraphFamily::Line) && !self.props.plot_area {
            return Err(Error::UnsupportedAuthoring {
                reason: "the standalone Graph profile requires a chart plot area",
            });
        }
        if family == GraphFamily::Pie && self.props.plot_area {
            return Err(Error::UnsupportedAuthoring {
                reason: "the standalone Graph pie profile has no AXES PlotArea record",
            });
        }
        if self.series.is_empty() {
            return Err(Error::InvalidModel {
                field: "series",
                reason: "a fresh Graph chart requires at least one series",
            });
        }
        if !self.formats.is_empty() || !self.labels.is_empty() || !self.unknown.is_empty() {
            return Err(Error::UnsupportedAuthoring {
                reason: "fresh Graph authoring does not synthesize opaque format or label records",
            });
        }
        let first_count = self
            .series
            .first()
            .map(|series| series.category_count)
            .ok_or(Error::InvalidModel {
                field: "series",
                reason: "a fresh Graph chart requires at least one series",
            })?;
        let category_count = usize::from(first_count.get());
        let expected_caches = self
            .series
            .len()
            .checked_mul(category_count.checked_add(1).ok_or(Error::SizeOverflow {
                resource: "Graph authoring cache count",
            })?)
            .and_then(|value| value.checked_add(category_count))
            .ok_or(Error::SizeOverflow {
                resource: "Graph authoring cache count",
            })?;
        if self.caches.len() != expected_caches {
            return Err(Error::InvalidModel {
                field: "cache",
                reason: "the Graph profile requires one name cell and one value cell per series point plus categories",
            });
        }
        let mut coordinates = Vec::new();
        coordinates
            .try_reserve_exact(self.caches.len())
            .ok()
            .ok_or(Error::Allocation {
                resource: "Graph authoring cache coordinates",
            })?;
        for cache in &self.caches {
            let Cache::Graph {
                row, col, value, ..
            } = cache
            else {
                return Err(Error::InvalidModel {
                    field: "cache",
                    reason: "fresh Graph authoring requires Graph cache cells",
                });
            };
            let row_index = usize::from(row.get());
            let col_index = usize::from(col.get());
            if row_index == 0 {
                if !(1..=category_count).contains(&col_index)
                    || !matches!(value, Value::Text(_) | Value::Blank)
                {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "the Graph category row must contain text or blank cells from column one",
                    });
                }
            } else if row_index > self.series.len()
                || col_index > category_count
                || (col_index == 0 && !matches!(value, Value::Text(_) | Value::Blank))
                || (col_index != 0 && !matches!(value, Value::Number(_) | Value::Blank))
            {
                return Err(Error::InvalidModel {
                    field: "cache",
                    reason: "the Graph series rows must contain numeric or blank cells from column zero",
                });
            }
            coordinates.push((row_index, col_index));
        }
        coordinates.sort_unstable();
        coordinates.dedup();
        if coordinates.len() != expected_caches {
            return Err(Error::InvalidModel {
                field: "cache",
                reason: "Graph authoring cache coordinates must form one complete datasheet grid",
            });
        }
        for (series_index, series) in self.series.iter().enumerate() {
            if series.category_count == Count::ZERO
                || series.value_count == Count::ZERO
                || series.category_count != series.value_count
                || series.category_count != first_count
                || series.bubble_count != Count::ZERO
                || series.owner != Owner::Group(GroupId::ZERO)
            {
                return Err(Error::InvalidModel {
                    field: "series",
                    reason: "the supported Graph profile requires equal non-empty category/value caches",
                });
            }
            for (binding, role) in series.ai.ordered().into_iter().zip(Role::ALL) {
                let Link::Graph { source, .. } = binding.link() else {
                    return Err(Error::InvalidModel {
                        field: "link",
                        reason: "fresh Graph series cannot contain Excel links",
                    });
                };
                let Link::Graph { row_col, .. } = binding.link() else {
                    unreachable!("Graph link was checked above");
                };
                let expected_row = match role {
                    Role::Categories | Role::Bubbles => RowCol::ZERO,
                    Role::Name | Role::Values => {
                        RowCol::new(u16::try_from(series_index + 1).ok().ok_or(
                            Error::SizeOverflow {
                                resource: "Graph authoring series row",
                            },
                        )?)
                        .ok_or(Error::InvalidModel {
                            field: "series",
                            reason: "Graph authoring series row exceeds the datasheet range",
                        })?
                    },
                };
                if *row_col != expected_row
                    || (role != Role::Bubbles && *source == Source::Automatic)
                    || (role == Role::Bubbles && *source != Source::Automatic)
                {
                    return Err(Error::InvalidModel {
                        field: "link",
                        reason: "fresh Graph series links do not match the canonical datasheet orientation",
                    });
                }
            }
            let name = series.ai.get(Role::Name).text();
            let name_cell = self.caches.iter().find(|cache| {
                matches!(
                    cache,
                    Cache::Graph { row, col, .. }
                        if row.get() == u16::try_from(series_index + 1).unwrap_or(u16::MAX)
                            && col.get() == 0
                )
            });
            let Some(Cache::Graph { value, .. }) = name_cell else {
                return Err(Error::InvalidModel {
                    field: "cache",
                    reason: "each Graph series name link requires its first datasheet cell",
                });
            };
            match (name, value) {
                (Some(expected), Value::Text(actual)) if expected == actual => {},
                (None, Value::Blank) => {},
                _ => {
                    return Err(Error::InvalidModel {
                        field: "cache",
                        reason: "the Graph series name cache disagrees with SeriesText",
                    });
                },
            }
        }
        codec::validate_chart(self, self.limits, true)
    }

    /// Parses a borrowed chart and retains an exact bounded copy for replay.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream is malformed or crosses its own limits;
    /// see [`Self::parse_with`].
    pub fn parse(input: Ref<'_>, context: Context) -> Result<Self> {
        let limits = input.limits();
        Self::parse_with(input, context, limits)
    }

    /// Parses a borrowed chart under explicit semantic limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] if the limits are inconsistent, or an
    /// error if the BIFF stream is malformed, crosses a configured bound, or
    /// cannot be copied within the allocation limits.
    pub fn parse_with(input: Ref<'_>, context: Context, limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        let mut chart = codec::parse(input, context, validated)?;
        chart.origin = Origin::Parsed(input.own_with(validated)?);
        Ok(chart)
    }

    /// Parses a move-owned stream without copying its input allocation.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream is malformed or crosses its own limits;
    /// see [`Self::open_with`].
    pub fn open(input: Stream, context: Context) -> Result<Self> {
        let limits = input.as_ref().limits();
        Self::open_with(input, context, limits)
    }

    /// Parses a move-owned stream under explicit semantic limits.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] if the limits are inconsistent, or an
    /// error if the stream cannot be rebounded, is malformed, or crosses a
    /// configured bound while parsing.
    pub fn open_with(input: Stream, context: Context, limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        let relimited = input.relimit(validated)?;
        let mut chart = codec::parse(relimited.as_ref(), context, validated)?;
        chart.origin = Origin::Parsed(relimited);
        Ok(chart)
    }

    /// Starts the bounded, lossless chart transaction.
    ///
    /// The transaction consumes this immutable chart snapshot. It changes
    /// only existing fixed-size cache, chart-area, or `ShtProps` records and
    /// keeps their producer-specific wire class, payload length, source
    /// identity, and record position. Unknown records, series identities,
    /// coordinates, formatting indices, and structural `PlotArea` presence
    /// therefore remain untouched.
    pub fn edit(self) -> Result<crate::chart::transaction::Editor> {
        crate::chart::transaction::Editor::new(self)
    }

    /// Replays an untouched parsed chart, or emits a proven fresh profile.
    pub fn encode(self) -> Result<Stream> {
        let limits = self.limits;
        self.encode_with(limits)
    }

    /// Consumes and replays under explicit bounds.
    ///
    /// Parsed mutations return [`Error::UnsafeEdit`]; fresh values outside the
    /// supported standalone Graph profile return [`Error::UnsupportedAuthoring`].
    pub fn encode_with(mut self, limits: Limits) -> Result<Stream> {
        let limits = limits.validate()?;
        let origin = std::mem::replace(&mut self.origin, Origin::Fresh);
        match origin {
            Origin::Parsed(_stream) if self.dirty => Err(Error::UnsafeEdit {
                reason: "opaque or reserved source records could not be placed losslessly",
            }),
            Origin::Parsed(stream) => stream.relimit(limits),
            Origin::Fresh => {
                if self.graph_authoring_profile {
                    // Public mutators can change a fresh model after the
                    // proof gate has been established. Re-run the complete
                    // profile proof immediately before emission so a stale
                    // internal flag cannot authorize an invalid stream.
                    self.validate_graph_authoring()?;
                }
                let bytes = codec::encode(&self, limits)?;
                Stream::with_limits(bytes, limits)
            },
        }
    }

    #[must_use]
    pub const fn context(&self) -> Context {
        self.context
    }

    #[must_use]
    pub const fn rect(&self) -> Rect {
        self.rect
    }

    pub fn set_rect(&mut self, value: Rect) {
        self.touch();
        self.rect = value;
    }

    #[must_use]
    pub const fn props(&self) -> Props {
        self.props
    }

    pub fn set_props(&mut self, value: Props) {
        self.touch();
        self.props = value;
    }

    /// Chart-window zoom.
    #[must_use]
    pub const fn zoom(&self) -> layout::Zoom {
        self.zoom
    }

    /// Sets the checked chart-window zoom.
    pub fn set_zoom(&mut self, value: layout::Zoom) {
        self.touch();
        self.zoom = value;
    }

    /// Plot-area font growth factors.
    #[must_use]
    pub const fn growth(&self) -> layout::Growth {
        self.growth
    }

    /// Sets plot-area font growth factors.
    pub fn set_growth(&mut self, value: layout::Growth) {
        self.touch();
        self.growth = value;
    }

    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_title(&mut self, value: Option<String>) {
        self.touch();
        self.title = value;
    }

    #[must_use]
    pub fn series(&self) -> &[Series] {
        &self.series
    }

    /// Mutably borrows series and marks parsed input dirty only on mutation.
    pub fn series_mut(&mut self) -> Edit<'_, Series> {
        Edit {
            values: &mut self.series,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_series(&mut self, value: Series) -> Result<()> {
        match &value.owner {
            Owner::Group(group) if usize::from(group.get()) >= self.groups.len() => {
                return Err(Error::InvalidModel {
                    field: "series",
                    reason: "series refers to a missing chart group",
                });
            },
            Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => {
                let zero_based = usize::try_from(parent.series().get() - 1).ok().ok_or({
                    Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index exceeds this platform",
                    }
                })?;
                if self
                    .series
                    .get(zero_based)
                    .is_none_or(|parent| !matches!(parent.owner, Owner::Group(_)))
                {
                    return Err(Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary series must refer to an existing regular series",
                    });
                }
            },
            Owner::Group(_) => {},
        }
        for binding in value.ai.ordered() {
            if !matches!(
                (self.context.kind(), binding.link()),
                (Kind::Excel, Link::Excel { .. }) | (Kind::Graph, Link::Graph { .. })
            ) {
                return Err(Error::InvalidModel {
                    field: "link",
                    reason: "series binding does not match the chart producer",
                });
            }
        }
        check_add(self.series.len(), self.limits.max_series, "series count")?;
        reserve_one(&mut self.series, "chart series")?;
        self.touch();
        self.series.push(value);
        Ok(())
    }

    /// Removes an unreferenced series and retargets later auxiliary parents.
    pub fn remove_series(&mut self, index: usize) -> Result<Option<Series>> {
        if index >= self.series.len() {
            return Ok(None);
        }
        let one_based = index.checked_add(1).ok_or(Error::SizeOverflow {
            resource: "series index",
        })?;
        let one_based = u32::try_from(one_based).ok().ok_or(Error::InvalidModel {
            field: "series",
            reason: "series index exceeds the auxiliary-parent range",
        })?;
        for series in &self.series {
            let parent = match &series.owner {
                Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => parent,
                Owner::Group(_) => continue,
            };
            let zero_based =
                usize::try_from(parent.series().get() - 1)
                    .ok()
                    .ok_or(Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index exceeds this platform",
                    })?;
            if self
                .series
                .get(zero_based)
                .is_none_or(|parent| !matches!(parent.owner, Owner::Group(_)))
            {
                return Err(Error::InvalidModel {
                    field: "series",
                    reason: "auxiliary series refers to an invalid parent",
                });
            }
        }
        if self.series.iter().any(|series| match &series.owner {
            Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => {
                parent.series().get() == one_based
            },
            Owner::Group(_) => false,
        }) {
            return Err(Error::InvalidModel {
                field: "series",
                reason: "series is still referenced by an auxiliary series",
            });
        }
        for series in &mut self.series {
            let parent = match &mut series.owner {
                Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => parent,
                Owner::Group(_) => continue,
            };
            if parent.series().get() > one_based {
                let shifted =
                    u16::try_from(parent.series().get() - 1)
                        .ok()
                        .ok_or(Error::InvalidModel {
                            field: "series",
                            reason: "auxiliary parent index exceeds its checked range",
                        })?;
                *parent = crate::record::series::Parent::try_new(shifted)
                    .ok()
                    .ok_or({
                        Error::InvalidModel {
                            field: "series",
                            reason: "auxiliary parent index became invalid",
                        }
                    })?;
            }
        }
        self.touch();
        Ok(Some(self.series.remove(index)))
    }

    #[must_use]
    pub fn groups(&self) -> &[Group] {
        &self.groups
    }

    /// Mutably borrows chart groups and marks parsed input dirty only on mutation.
    pub fn groups_mut(&mut self) -> Edit<'_, Group> {
        Edit {
            values: &mut self.groups,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_group(&mut self, value: Group) -> Result<()> {
        check_add(self.groups.len(), self.limits.max_groups, "group count")?;
        if self
            .parents
            .get(value.parent.index())
            .is_none_or(|parent| parent.id() != value.parent)
        {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart group refers to a missing axis parent",
            });
        }
        if self.groups.iter().any(|group| group.order == value.order) {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart-group drawing order is duplicated",
            });
        }
        reserve_one(&mut self.groups, "chart groups")?;
        self.touch();
        self.groups.push(value);
        Ok(())
    }

    /// Removes an unreferenced group and retargets later group indices.
    ///
    /// A referenced group is refused instead of silently moving its series to
    /// a different chart family.
    pub fn remove_group(&mut self, index: usize) -> Result<Option<Group>> {
        if index >= self.groups.len() {
            return Ok(None);
        }
        let raw = u8::try_from(index).ok().ok_or(Error::InvalidModel {
            field: "group",
            reason: "chart-group index exceeds nine",
        })?;
        let id = GroupId::new(raw).ok_or(Error::InvalidModel {
            field: "group",
            reason: "chart-group index exceeds nine",
        })?;
        if self.series.iter().any(|series| {
            series
                .owner
                .group()
                .is_some_and(|group| usize::from(group.get()) >= self.groups.len())
        }) {
            return Err(Error::InvalidModel {
                field: "series",
                reason: "series refers to an invalid chart group",
            });
        }
        if self
            .series
            .iter()
            .any(|series| series.owner.group() == Some(id))
        {
            return Err(Error::InvalidModel {
                field: "group",
                reason: "chart group is still referenced by a series",
            });
        }
        for series in &mut self.series {
            if let Owner::Group(group) = &mut series.owner
                && group.get() > raw
            {
                *group = GroupId::new(group.get() - 1).ok_or(Error::InvalidModel {
                    field: "series",
                    reason: "series chart-group index became invalid",
                })?;
            }
        }
        self.touch();
        Ok(Some(self.groups.remove(index)))
    }

    #[must_use]
    pub fn axes(&self) -> &[axis::Axis] {
        &self.axes
    }

    /// Mutably borrows axes and marks parsed input dirty only on mutation.
    pub fn axes_mut(&mut self) -> Edit<'_, axis::Axis> {
        Edit {
            values: &mut self.axes,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_axis(&mut self, value: axis::Axis) -> Result<()> {
        if self
            .parents
            .get(value.parent.index())
            .is_none_or(|parent| parent.id() != value.parent)
        {
            return Err(Error::InvalidModel {
                field: "axis",
                reason: "axis refers to a missing axis parent",
            });
        }
        check_add(self.axes.len(), self.limits.max_axes, "axis count")?;
        reserve_one(&mut self.axes, "chart axes")?;
        self.touch();
        self.axes.push(value);
        Ok(())
    }

    pub fn remove_axis(&mut self, index: usize) -> Option<axis::Axis> {
        if index >= self.axes.len() {
            return None;
        }
        self.touch();
        Some(self.axes.remove(index))
    }

    /// Primary and optional secondary axis-parent collections.
    #[must_use]
    pub fn parents(&self) -> &[axis::Parent] {
        &self.parents
    }

    /// Mutably borrows axis-parent metadata.
    pub fn parents_mut(&mut self) -> Edit<'_, axis::Parent> {
        Edit {
            values: &mut self.parents,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    #[must_use]
    pub const fn legend(&self) -> Option<Legend> {
        self.legend
    }

    pub fn set_legend(&mut self, value: Option<Legend>) {
        self.touch();
        self.legend = value;
    }

    #[must_use]
    pub fn caches(&self) -> &[Cache] {
        &self.caches
    }

    /// Context-specific mandatory cache dimensions.
    #[must_use]
    pub const fn dimensions(&self) -> chart_cache::Dims {
        self.dimensions
    }

    /// Sets producer-typed cache dimensions.
    pub fn set_dimensions(&mut self, value: chart_cache::Dims) -> Result<()> {
        if !value.matches(self.context.kind()) {
            return Err(Error::InvalidModel {
                field: "Dimensions",
                reason: "dimensions do not match the chart producer",
            });
        }
        let derived = cache_dimensions(&self.caches, self.context.kind())?;
        if !dimensions_cover(value, derived) {
            return Err(Error::InvalidModel {
                field: "Dimensions",
                reason: "dimensions do not cover the cached chart cells",
            });
        }
        self.touch();
        self.dimensions = value;
        Ok(())
    }

    /// Mutably borrows cached cells and marks parsed input dirty only on mutation.
    /// Call [`Self::sync_dimensions`] after changing cell coordinates.
    pub fn caches_mut(&mut self) -> Edit<'_, Cache> {
        Edit {
            values: &mut self.caches,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_cache(&mut self, value: Cache) -> Result<()> {
        if value.kind() != self.context.kind() {
            return Err(Error::InvalidModel {
                field: "cache",
                reason: "cached cell does not match the chart producer",
            });
        }
        check_add(
            self.caches.len(),
            self.limits.max_cached_values,
            "cached value count",
        )?;
        reserve_one(&mut self.caches, "chart cache")?;
        self.caches.push(value);
        let dimensions = match cache_dimensions(&self.caches, self.context.kind()) {
            Ok(value) => value,
            Err(error) => {
                let _ = self.caches.pop();
                return Err(error);
            },
        };
        self.touch();
        self.dimensions = dimensions;
        Ok(())
    }

    /// Removes one cached cell and synchronizes producer dimensions.
    pub fn remove_cache(&mut self, index: usize) -> Result<Option<Cache>> {
        if index >= self.caches.len() {
            return Ok(None);
        }
        let removed = self.caches.remove(index);
        let dimensions = match cache_dimensions(&self.caches, self.context.kind()) {
            Ok(value) => value,
            Err(error) => {
                self.caches.insert(index, removed);
                return Err(error);
            },
        };
        self.touch();
        self.dimensions = dimensions;
        Ok(Some(removed))
    }

    /// Recomputes mandatory dimensions from the current cached cells.
    pub fn sync_dimensions(&mut self) -> Result<()> {
        let dimensions = cache_dimensions(&self.caches, self.context.kind())?;
        self.touch();
        self.dimensions = dimensions;
        Ok(())
    }

    #[must_use]
    pub fn formats(&self) -> &[format::Format] {
        &self.formats
    }

    pub fn formats_mut(&mut self) -> Edit<'_, format::Format> {
        Edit {
            values: &mut self.formats,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_format(&mut self, value: format::Format) -> Result<()> {
        check_add(
            self.formats.len(),
            self.limits.max_chart_records,
            "chart record count",
        )?;
        reserve_one(&mut self.formats, "chart formats")?;
        self.touch();
        self.formats.push(value);
        Ok(())
    }

    #[must_use]
    pub fn labels(&self) -> &[Label] {
        &self.labels
    }

    pub fn labels_mut(&mut self) -> Edit<'_, Label> {
        Edit {
            values: &mut self.labels,
            dirty: &mut self.dirty,
            parsed: matches!(&self.origin, Origin::Parsed(_)),
        }
    }

    pub fn add_label(&mut self, value: Label) -> Result<()> {
        check_add(
            self.labels.len(),
            self.limits.max_chart_records,
            "chart record count",
        )?;
        reserve_one(&mut self.labels, "chart labels")?;
        self.touch();
        self.labels.push(value);
        Ok(())
    }

    /// Unknown and recognized-but-opaque records in original encounter order.
    #[must_use]
    pub fn unknown(&self) -> &[Raw] {
        &self.unknown
    }

    /// Whether this parsed chart still has an exact replayable source stream.
    #[must_use]
    pub fn is_pristine(&self) -> bool {
        matches!(&self.origin, Origin::Parsed(_)) && !self.dirty
    }

    fn touch(&mut self) {
        if matches!(&self.origin, Origin::Parsed(_)) {
            self.dirty = true;
        }
    }
}
