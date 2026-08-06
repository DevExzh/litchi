use super::super::super::{Kind, Ref, Stream, axis, cache as chart_cache, codec, format, layout};
use super::super::cache::Cache;
use super::super::context::{Context, GroupId, Props, Rect};
use super::super::groups::Group;
use super::super::inventory::{Edit, Label, Legend, Origin, Raw};
use super::super::series::{Link, Owner, Series};
use super::validation::{cache_dimensions, check_add, dimensions_cover, reserve_one};
use crate::{Error, Limits, Result};

/// Host-neutral semantic chart.
///
/// Parsed values retain their exact [`Stream`]. An untouched parsed value can
/// therefore be encoded byte-for-byte without copying. Mutation is allowed for
/// inspection workflows, but encoding such a value is refused until a future
/// lossless record editor can prove placement of every opaque record.
/// Fresh values can be assembled and validated, but encoding currently returns
/// [`Error::UnsupportedAuthoring`] until the complete mandatory chart-sheet
/// scaffold is represented. This prevents self-consistent but Office-invalid
/// streams from escaping the crate.
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
    /// Internal proof gate. No public constructor enables it until the full
    /// CHARTSHEET/CHARTFOMATS/SERIESDATA grammar is modeled.
    pub(in crate::chart) authoring_proven: bool,
}

impl Chart {
    /// Creates a fresh chart using conservative limits.
    pub fn new(context: Context) -> Result<Self> {
        Self::new_with(context, Limits::default())
    }

    /// Creates a fresh chart with explicit authoring bounds.
    pub fn new_with(context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut groups = Vec::new();
        groups.try_reserve_exact(1).map_err(|_| Error::Allocation {
            resource: "chart groups",
        })?;
        groups.push(Group::line());
        let mut parents = Vec::new();
        parents
            .try_reserve_exact(1)
            .map_err(|_| Error::Allocation {
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
            limits,
            authoring_proven: false,
        })
    }

    /// Parses a borrowed chart and retains an exact bounded copy for replay.
    pub fn parse(input: Ref<'_>, context: Context) -> Result<Self> {
        let limits = input.limits();
        Self::parse_with(input, context, limits)
    }

    /// Parses a borrowed chart under explicit semantic limits.
    pub fn parse_with(input: Ref<'_>, context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let mut chart = codec::parse(input, context, limits)?;
        chart.origin = Origin::Parsed(input.own_with(limits)?);
        Ok(chart)
    }

    /// Parses a move-owned stream without copying its input allocation.
    pub fn open(input: Stream, context: Context) -> Result<Self> {
        let limits = input.as_ref().limits();
        Self::open_with(input, context, limits)
    }

    /// Parses a move-owned stream under explicit semantic limits.
    pub fn open_with(input: Stream, context: Context, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let input = input.relimit(limits)?;
        let mut chart = codec::parse(input.as_ref(), context, limits)?;
        chart.origin = Origin::Parsed(input);
        Ok(chart)
    }

    /// Replays an untouched parsed chart, or refuses unsupported fresh authoring.
    pub fn encode(self) -> Result<Stream> {
        let limits = self.limits;
        self.encode_with(limits)
    }

    /// Consumes and replays under explicit bounds.
    ///
    /// Parsed mutations return [`Error::UnsafeEdit`]; fresh values return
    /// [`Error::UnsupportedAuthoring`] until full chart-sheet authoring lands.
    pub fn encode_with(mut self, limits: Limits) -> Result<Stream> {
        let limits = limits.validate()?;
        let origin = std::mem::replace(&mut self.origin, Origin::Fresh);
        match origin {
            Origin::Parsed(_stream) if self.dirty => Err(Error::UnsafeEdit {
                reason: "opaque or reserved source records could not be placed losslessly",
            }),
            Origin::Parsed(stream) => stream.relimit(limits),
            Origin::Fresh => {
                let bytes = codec::encode(&self, limits)?;
                Stream::with_limits(bytes, limits)
            },
        }
    }

    pub const fn context(&self) -> Context {
        self.context
    }

    pub const fn rect(&self) -> Rect {
        self.rect
    }

    pub fn set_rect(&mut self, value: Rect) {
        self.touch();
        self.rect = value;
    }

    pub const fn props(&self) -> Props {
        self.props
    }

    pub fn set_props(&mut self, value: Props) {
        self.touch();
        self.props = value;
    }

    /// Chart-window zoom.
    pub const fn zoom(&self) -> layout::Zoom {
        self.zoom
    }

    /// Sets the checked chart-window zoom.
    pub fn set_zoom(&mut self, value: layout::Zoom) {
        self.touch();
        self.zoom = value;
    }

    /// Plot-area font growth factors.
    pub const fn growth(&self) -> layout::Growth {
        self.growth
    }

    /// Sets plot-area font growth factors.
    pub fn set_growth(&mut self, value: layout::Growth) {
        self.touch();
        self.growth = value;
    }

    pub fn title(&self) -> Option<&str> {
        self.title.as_deref()
    }

    pub fn set_title(&mut self, value: Option<String>) {
        self.touch();
        self.title = value;
    }

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
                let zero_based = usize::try_from(parent.series().get() - 1).map_err(|_| {
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
            _ => {},
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
        let one_based = u32::try_from(one_based).map_err(|_| Error::InvalidModel {
            field: "series",
            reason: "series index exceeds the auxiliary-parent range",
        })?;
        for series in &self.series {
            let parent = match &series.owner {
                Owner::Trend { parent, .. } | Owner::ErrorBar { parent, .. } => parent,
                Owner::Group(_) => continue,
            };
            let zero_based =
                usize::try_from(parent.series().get() - 1).map_err(|_| Error::InvalidModel {
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
                    u16::try_from(parent.series().get() - 1).map_err(|_| Error::InvalidModel {
                        field: "series",
                        reason: "auxiliary parent index exceeds its checked range",
                    })?;
                *parent = crate::record::series::Parent::try_new(shifted).map_err(|_| {
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
        let raw = u8::try_from(index).map_err(|_| Error::InvalidModel {
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

    pub const fn legend(&self) -> Option<Legend> {
        self.legend
    }

    pub fn set_legend(&mut self, value: Option<Legend>) {
        self.touch();
        self.legend = value;
    }

    pub fn caches(&self) -> &[Cache] {
        &self.caches
    }

    /// Context-specific mandatory cache dimensions.
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
    pub fn unknown(&self) -> &[Raw] {
        &self.unknown
    }

    /// Whether this parsed chart still has an exact replayable source stream.
    pub fn is_pristine(&self) -> bool {
        matches!(&self.origin, Origin::Parsed(_)) && !self.dirty
    }

    fn touch(&mut self) {
        if matches!(&self.origin, Origin::Parsed(_)) {
            self.dirty = true;
        }
    }
}
