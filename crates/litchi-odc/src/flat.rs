//! Bounded, byte-preserving flat `OpenDocument` Chart snapshots and axis edits.

use litchi_core::{Error, FileFormat, Result};
use litchi_odf_common::chart::{self, Element, PlotArea};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{collections::BTreeSet, ops::Range, sync::Arc};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const CHART: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
const SVG: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const MAX_NAME_BYTES: usize = 64 * 1024;

#[derive(Clone, Debug)]
struct State {
    bytes: Vec<u8>,
    chart: Element,
    axes: Vec<AxisRecord>,
    chart_tag: EditableTag,
    plot_tag: EditableTag,
    series_tags: Vec<EditableTag>,
    root_kind: RootKind,
    limits: crate::Limits,
}

#[derive(Clone, Debug)]
struct EditableTag {
    tag: Range<usize>,
    prefix: Option<String>,
    attributes: Vec<AttributeRecord>,
}

#[derive(Clone, Debug)]
struct AttributeRecord {
    attribute: ExactAttribute,
    value: String,
    value_span: Range<usize>,
    attribute_span: Range<usize>,
}

#[derive(Clone, Debug)]
struct AxisRecord {
    name: Option<String>,
    style_name: Option<String>,
    tag: Range<usize>,
    name_value: Option<Range<usize>>,
    name_attribute: Option<Range<usize>>,
    style_value: Option<Range<usize>>,
    style_attribute: Option<Range<usize>>,
    prefix: Option<String>,
}

/// An immutable, byte-exact flat ODC snapshot.
#[derive(Clone, Debug)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the public name distinguishes a flat chart from a packaged chart"
)]
pub struct FlatChart(Arc<State>);

impl FlatChart {
    /// Opens a flat ODC document from owned UTF-8 XML bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the input is not a bounded, structurally valid
    /// flat ODC document.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, crate::Limits::default())
    }

    /// Opens flat ODC XML under caller-selected limits retained by edits.
    ///
    /// # Errors
    ///
    /// Returns an error when XML structure, semantics, or a limit is invalid.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: crate::Limits) -> Result<Self> {
        parse(bytes, RootKind::Flat, limits).map(|state| Self(Arc::new(state)))
    }

    pub(crate) fn from_content_xml(bytes: Vec<u8>, limits: crate::Limits) -> Result<Self> {
        parse(bytes, RootKind::Content, limits).map(|state| Self(Arc::new(state)))
    }

    /// Returns the retained semantic chart tree.
    #[must_use]
    pub fn chart(&self) -> &Element {
        &self.0.chart
    }

    /// Returns the first chart plot area, when present.
    #[must_use]
    pub fn plot_area(&self) -> Option<PlotArea<'_>> {
        self.chart().plot_area()
    }

    /// Finds the first axis with the requested `chart:name`.
    #[must_use]
    pub fn find_axis(&self, name: &str) -> Option<chart::Axis<'_>> {
        self.plot_area()?
            .axes()
            .find(|axis| axis.name() == Some(name))
    }

    /// Returns the exact original XML bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0.bytes
    }

    /// Return the limits retained for subsequent edits and patch application.
    #[must_use]
    pub fn limits(&self) -> crate::Limits {
        self.0.limits
    }

    /// Return one controlled attribute value after namespace resolution.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or target/attribute pairing.
    pub fn exact_value(
        &self,
        target: ExactTarget,
        attribute: ExactAttribute,
    ) -> Result<Option<&str>> {
        validate_exact_pair(target, attribute)?;
        Ok(self.tag(target)?.value(attribute))
    }

    /// Starts a detached transaction bound to this source snapshot.
    #[must_use]
    pub fn edit(&self) -> FlatChartEdit {
        FlatChartEdit {
            source: self.clone(),
            changes: Vec::new(),
            exact_changes: Vec::new(),
        }
    }

    /// Consumes the snapshot and returns its exact XML bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }

    fn axis_name(&self, index: usize) -> Result<Option<&str>> {
        self.0
            .axes
            .get(index)
            .map(|axis| axis.name.as_deref())
            .ok_or_else(|| invalid_error("flat ODC axis selector is out of bounds"))
    }

    fn tag(&self, target: ExactTarget) -> Result<&EditableTag> {
        match target {
            ExactTarget::Chart => Ok(&self.0.chart_tag),
            ExactTarget::PlotArea => Ok(&self.0.plot_tag),
            ExactTarget::Series(index) => self
                .0
                .series_tags
                .get(index)
                .ok_or_else(|| invalid_error("flat ODC series selector is out of bounds")),
        }
    }
}

impl EditableTag {
    fn value(&self, attribute: ExactAttribute) -> Option<&str> {
        self.attributes
            .iter()
            .find(|record| record.attribute == attribute)
            .map(|record| record.value.as_str())
    }

    fn record(&self, attribute: ExactAttribute) -> Option<&AttributeRecord> {
        self.attributes
            .iter()
            .find(|record| record.attribute == attribute)
    }
}

/// A byte-preserving exact-edit target outside the axis compatibility surface.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum ExactTarget {
    /// The root `chart:chart` element.
    Chart,
    /// The chart's single `chart:plot-area` element.
    PlotArea,
    /// A direct plot-area series selected by zero-based position.
    Series(usize),
}

/// A controlled chart attribute supported by exact-span editing.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum ExactAttribute {
    Class,
    StyleName,
    Width,
    Height,
    X,
    Y,
    CellRangeAddress,
    ValuesCellRangeAddress,
    LabelCellAddress,
    AttachedAxis,
}

/// One validated exact-span attribute change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ExactChange {
    target: ExactTarget,
    attribute: ExactAttribute,
    before: Option<String>,
    after: Option<String>,
}

impl ExactChange {
    pub(crate) fn new_inverse(change: &Self) -> Self {
        Self {
            target: change.target,
            attribute: change.attribute,
            before: change.after.clone(),
            after: change.before.clone(),
        }
    }

    #[must_use]
    pub const fn target(&self) -> ExactTarget {
        self.target
    }

    #[must_use]
    pub const fn attribute(&self) -> ExactAttribute {
        self.attribute
    }

    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }
}

/// A partial update for one chart axis.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct AxisUpdate {
    /// `None` leaves the name unchanged, `Some(None)` removes it, and
    /// `Some(Some(name))` sets it.
    pub name: Option<Option<String>>,
    /// `None` leaves the style unchanged, while `Some` sets or removes it.
    pub style_name: Option<Option<String>>,
}

impl AxisUpdate {
    /// Creates an update that assigns `chart:name`.
    #[must_use]
    pub fn named(name: impl Into<String>) -> Self {
        Self {
            name: Some(Some(name.into())),
            style_name: None,
        }
    }

    /// Creates an update that removes `chart:name`.
    #[must_use]
    pub const fn unnamed() -> Self {
        Self {
            name: Some(None),
            style_name: None,
        }
    }

    /// Creates an update that sets `chart:style-name` exactly in place.
    #[must_use]
    pub fn styled(style_name: impl Into<String>) -> Self {
        Self {
            name: None,
            style_name: Some(Some(style_name.into())),
        }
    }

    /// Creates an update that removes `chart:style-name`.
    #[must_use]
    pub const fn unstyled() -> Self {
        Self {
            name: None,
            style_name: Some(None),
        }
    }
}

/// A staged, source-bound flat chart transaction.
#[allow(
    clippy::module_name_repetitions,
    reason = "the public name associates this transaction with FlatChart"
)]
pub struct FlatChartEdit {
    source: FlatChart,
    changes: Vec<AxisChange>,
    exact_changes: Vec<ExactChange>,
}

impl FlatChartEdit {
    /// Stage a controlled attribute edit on the chart, plot area, or a series.
    ///
    /// `None` removes an existing optional attribute. Adding an absent
    /// attribute is supported only for chart-namespace attributes whose
    /// namespace prefix is proven by the selected element name.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid target/attribute pair, selector, value,
    /// or a namespace insertion that cannot be represented losslessly.
    pub fn update_exact(
        &mut self,
        target: ExactTarget,
        attribute: ExactAttribute,
        after: Option<String>,
    ) -> Result<()> {
        validate_exact_pair(target, attribute)?;
        if after.as_ref().is_some_and(|value| {
            value.len() > MAX_NAME_BYTES || value.len() > self.source.0.limits.max_scalar_bytes()
        }) {
            return Err(invalid_error("flat ODC attribute exceeds its byte limit"));
        }
        let tag = self.source.tag(target)?;
        let before = tag.value(attribute).map(str::to_owned);
        let current = self
            .exact_changes
            .iter()
            .find(|change| change.target == target && change.attribute == attribute)
            .map_or_else(|| before.clone(), |change| change.after.clone());
        if current == after {
            return Ok(());
        }
        if before == after {
            self.exact_changes
                .retain(|change| change.target != target || change.attribute != attribute);
            return Ok(());
        }
        let change = ExactChange {
            target,
            attribute,
            before,
            after,
        };
        if let Some(slot) = self
            .exact_changes
            .iter_mut()
            .find(|candidate| candidate.target == target && candidate.attribute == attribute)
        {
            *slot = change;
        } else {
            self.exact_changes.push(change);
        }
        Ok(())
    }

    /// Stages an axis update by plot-area axis index.
    ///
    /// # Errors
    ///
    /// Returns an error when the selector is out of bounds or the requested
    /// name exceeds the bounded value limit.
    pub fn update_axis(&mut self, index: usize, update: AxisUpdate) -> Result<()> {
        let axis = self
            .source
            .0
            .axes
            .get(index)
            .ok_or_else(|| invalid_error("flat ODC axis selector is out of bounds"))?;
        let staged_change = self.changes.iter().find(|change| change.index == index);
        let current_name =
            staged_change.map_or_else(|| axis.name.clone(), |change| change.after.clone());
        let current_style = staged_change.map_or_else(
            || axis.style_name.clone(),
            |change| change.after_style_name.clone(),
        );
        let after = update.name.unwrap_or(current_name);
        let after_style = update.style_name.unwrap_or(current_style);
        if after.as_ref().is_some_and(|name| {
            name.len() > MAX_NAME_BYTES || name.len() > self.source.0.limits.max_scalar_bytes()
        }) || after_style.as_ref().is_some_and(|name| {
            name.len() > MAX_NAME_BYTES || name.len() > self.source.0.limits.max_scalar_bytes()
        }) {
            return Err(invalid_error("flat ODC axis name exceeds its byte limit"));
        }
        let before = self.source.axis_name(index)?.map(str::to_owned);
        let before_style = axis.style_name.clone();
        if before == after && before_style == after_style {
            self.changes.retain(|change| change.index != index);
            return Ok(());
        }
        let change = AxisChange {
            index,
            before,
            after,
            before_style_name: before_style,
            after_style_name: after_style,
        };
        if let Some(change_slot) = self
            .changes
            .iter_mut()
            .find(|candidate| candidate.index == index)
        {
            *change_slot = change;
        } else {
            self.changes.push(change);
        }
        Ok(())
    }

    /// Validates and atomically publishes all staged changes in memory.
    ///
    /// # Errors
    ///
    /// Returns an error when a change cannot be applied losslessly or the
    /// candidate fails structural validation and typed readback.
    pub fn commit(self) -> Result<FlatChartCommit> {
        let FlatChartEdit {
            source,
            changes,
            exact_changes,
        } = self;
        let mut replacements = Vec::with_capacity(changes.len() + exact_changes.len());
        for change in &changes {
            let axis = &source.0.axes[change.index];
            if change.before != change.after {
                stage_axis_attribute(
                    &source,
                    axis,
                    "name",
                    change.after.as_ref(),
                    axis.name_value.as_ref(),
                    axis.name_attribute.as_ref(),
                    &mut replacements,
                )?;
            }
            if change.before_style_name != change.after_style_name {
                stage_axis_attribute(
                    &source,
                    axis,
                    "style-name",
                    change.after_style_name.as_ref(),
                    axis.style_value.as_ref(),
                    axis.style_attribute.as_ref(),
                    &mut replacements,
                )?;
            }
        }
        for change in &exact_changes {
            let tag = source.tag(change.target)?;
            stage_exact_attribute(
                &source,
                tag,
                change.attribute,
                change.after.as_ref(),
                &mut replacements,
            )?;
        }
        replacements.sort_unstable_by_key(|replacement| std::cmp::Reverse(replacement.range.start));
        let mut bytes = source.as_bytes().to_vec();
        let mut previous = bytes.len();
        for replacement in replacements {
            if replacement.range.end > previous
                || replacement.range.start > replacement.range.end
                || replacement.range.end > bytes.len()
            {
                return Err(invalid_error("flat ODC edit contains overlapping spans"));
            }
            bytes.splice(replacement.range.clone(), replacement.value);
            previous = replacement.range.start;
        }
        let snapshot = FlatChart(Arc::new(parse(bytes, source.0.root_kind, source.0.limits)?));
        for change in &changes {
            if snapshot.axis_name(change.index)? != change.after.as_deref() {
                return Err(invalid_error("flat ODC edit failed typed readback"));
            }
            if snapshot.0.axes[change.index].style_name.as_deref()
                != change.after_style_name.as_deref()
            {
                return Err(invalid_error("flat ODC style edit failed typed readback"));
            }
        }
        for change in &exact_changes {
            if snapshot.tag(change.target)?.value(change.attribute) != change.after.as_deref() {
                return Err(invalid_error("flat ODC exact edit failed typed readback"));
            }
        }
        Ok(FlatChartCommit {
            snapshot: snapshot.clone(),
            patch: FlatChartPatch {
                source,
                target: snapshot,
                changes,
                exact_changes,
            },
        })
    }
}

/// A committed flat chart snapshot and its reversible semantic patch.
#[allow(
    clippy::module_name_repetitions,
    reason = "the public name associates this commit with FlatChart"
)]
pub struct FlatChartCommit {
    snapshot: FlatChart,
    patch: FlatChartPatch,
}

impl FlatChartCommit {
    /// Returns the committed snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &FlatChart {
        &self.snapshot
    }

    /// Returns the source-bound patch.
    #[must_use]
    pub fn patch(&self) -> &FlatChartPatch {
        &self.patch
    }

    /// Consumes the commit and returns its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> FlatChart {
        self.snapshot
    }
}

/// A source-checked, reversible collection of axis changes.
#[derive(Clone, Debug)]
#[allow(
    clippy::module_name_repetitions,
    reason = "the public name associates this patch with FlatChart"
)]
pub struct FlatChartPatch {
    source: FlatChart,
    target: FlatChart,
    changes: Vec<AxisChange>,
    exact_changes: Vec<ExactChange>,
}

impl PartialEq for FlatChartPatch {
    fn eq(&self, other: &Self) -> bool {
        self.source.as_bytes() == other.source.as_bytes()
            && self.target.as_bytes() == other.target.as_bytes()
            && self.changes == other.changes
            && self.exact_changes == other.exact_changes
    }
}

impl Eq for FlatChartPatch {}

impl FlatChartPatch {
    /// Returns whether this patch was committed from the exact supplied XML.
    #[must_use]
    pub fn is_applicable_to(&self, source: &FlatChart) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Returns the ordered semantic changes.
    #[must_use]
    pub fn changes(&self) -> &[AxisChange] {
        &self.changes
    }

    /// Returns the ordered chart, plot-area, and series attribute changes.
    #[must_use]
    pub fn exact_changes(&self) -> &[ExactChange] {
        &self.exact_changes
    }

    pub(crate) fn tag_splices(&self) -> Result<Vec<TagSplice>> {
        if self.source.0.axes.len() != self.target.0.axes.len() {
            return Err(invalid_error(
                "flat ODC axis splice changed the axis inventory",
            ));
        }
        if self.source.0.series_tags.len() != self.target.0.series_tags.len() {
            return Err(invalid_error(
                "flat ODC splice changed the series inventory",
            ));
        }
        let mut targets = self
            .changes
            .iter()
            .map(|change| SpliceTarget::Axis(change.index))
            .chain(
                self.exact_changes
                    .iter()
                    .map(|change| SpliceTarget::Exact(change.target)),
            )
            .collect::<BTreeSet<_>>();
        let mut splices = Vec::with_capacity(targets.len());
        for splice_target in std::mem::take(&mut targets) {
            let (source_tag, target_tag) = match splice_target {
                SpliceTarget::Axis(index) => (
                    self.source.0.axes[index].tag.clone(),
                    self.target.0.axes[index].tag.clone(),
                ),
                SpliceTarget::Exact(exact_target) => (
                    self.source.tag(exact_target)?.tag.clone(),
                    self.target.tag(exact_target)?.tag.clone(),
                ),
            };
            splices.push(TagSplice {
                range: source_tag.clone(),
                expected: self.source.as_bytes()[source_tag].to_vec(),
                replacement: self.target.as_bytes()[target_tag].to_vec(),
            });
        }
        let mut rebuilt = self.source.as_bytes().to_vec();
        splices.sort_unstable_by_key(|splice| std::cmp::Reverse(splice.range.start));
        for splice in &splices {
            rebuilt.splice(splice.range.clone(), splice.replacement.iter().copied());
        }
        if rebuilt != self.target.as_bytes() {
            return Err(invalid_error(
                "flat ODC tag splice does not reproduce its committed target",
            ));
        }
        Ok(splices)
    }

    pub(crate) fn target_bytes(&self) -> &[u8] {
        self.target.as_bytes()
    }

    pub(crate) fn source_bytes(&self) -> &[u8] {
        self.source.as_bytes()
    }

    /// Returns a patch that reverses this patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .map(|change| AxisChange {
                    index: change.index,
                    before: change.after.clone(),
                    after: change.before.clone(),
                    before_style_name: change.after_style_name.clone(),
                    after_style_name: change.before_style_name.clone(),
                })
                .collect(),
            exact_changes: self
                .exact_changes
                .iter()
                .map(|change| ExactChange {
                    target: change.target,
                    attribute: change.attribute,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
        }
    }

    /// Applies the patch only to its exact immutable source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not the exact snapshot from which
    /// this patch was committed.
    pub fn apply(&self, source: &FlatChart) -> Result<FlatChartCommit> {
        if !self.is_applicable_to(source) {
            return Err(invalid_error("flat ODC patch source does not match"));
        }
        Ok(FlatChartCommit {
            snapshot: self.target.clone(),
            patch: self.clone(),
        })
    }
}

/// One selector-bound axis name and style-reference change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct AxisChange {
    index: usize,
    before: Option<String>,
    after: Option<String>,
    before_style_name: Option<String>,
    after_style_name: Option<String>,
}

impl AxisChange {
    pub(crate) fn new(index: usize, before: Option<String>, after: Option<String>) -> Self {
        Self {
            index,
            before,
            after,
            before_style_name: None,
            after_style_name: None,
        }
    }

    pub(crate) fn with_style(
        mut self,
        before_style_name: Option<String>,
        after_style_name: Option<String>,
    ) -> Self {
        self.before_style_name = before_style_name;
        self.after_style_name = after_style_name;
        self
    }

    pub(crate) fn new_inverse(change: &Self) -> Self {
        Self {
            index: change.index,
            before: change.after.clone(),
            after: change.before.clone(),
            before_style_name: change.after_style_name.clone(),
            after_style_name: change.before_style_name.clone(),
        }
    }

    #[must_use]
    pub fn index(&self) -> usize {
        self.index
    }

    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Return the style reference before the edit.
    #[must_use]
    pub fn before_style_name(&self) -> Option<&str> {
        self.before_style_name.as_deref()
    }

    /// Return the style reference after the edit.
    #[must_use]
    pub fn after_style_name(&self) -> Option<&str> {
        self.after_style_name.as_deref()
    }
}

struct Replacement {
    range: Range<usize>,
    value: Vec<u8>,
}

pub(crate) struct TagSplice {
    pub(crate) range: Range<usize>,
    pub(crate) expected: Vec<u8>,
    pub(crate) replacement: Vec<u8>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
enum SpliceTarget {
    Axis(usize),
    Exact(ExactTarget),
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RootKind {
    Flat,
    Content,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Chart,
    Other,
}

pub(crate) fn exact_changes_between(source: &FlatChart, target: &FlatChart) -> Vec<ExactChange> {
    if source.0.series_tags.len() != target.0.series_tags.len() {
        return Vec::new();
    }
    let targets = std::iter::once(ExactTarget::Chart)
        .chain(std::iter::once(ExactTarget::PlotArea))
        .chain((0..source.0.series_tags.len()).map(ExactTarget::Series));
    let mut changes = Vec::new();
    for exact_target in targets {
        let Ok(source_tag) = source.tag(exact_target) else {
            continue;
        };
        let Ok(target_tag) = target.tag(exact_target) else {
            continue;
        };
        let attributes = source_tag
            .attributes
            .iter()
            .map(|record| record.attribute)
            .chain(target_tag.attributes.iter().map(|record| record.attribute))
            .collect::<BTreeSet<_>>();
        for attribute in attributes {
            let before = source_tag.value(attribute).map(str::to_owned);
            let after = target_tag.value(attribute).map(str::to_owned);
            if before != after {
                changes.push(ExactChange {
                    target: exact_target,
                    attribute,
                    before,
                    after,
                });
            }
        }
    }
    changes
}

fn validate_exact_pair(target: ExactTarget, attribute: ExactAttribute) -> Result<()> {
    let valid = match target {
        ExactTarget::Chart => matches!(
            attribute,
            ExactAttribute::Class
                | ExactAttribute::StyleName
                | ExactAttribute::Width
                | ExactAttribute::Height
        ),
        ExactTarget::PlotArea => matches!(
            attribute,
            ExactAttribute::StyleName
                | ExactAttribute::CellRangeAddress
                | ExactAttribute::X
                | ExactAttribute::Y
                | ExactAttribute::Width
                | ExactAttribute::Height
        ),
        ExactTarget::Series(_) => matches!(
            attribute,
            ExactAttribute::Class
                | ExactAttribute::StyleName
                | ExactAttribute::ValuesCellRangeAddress
                | ExactAttribute::LabelCellAddress
                | ExactAttribute::AttachedAxis
        ),
    };
    if valid {
        Ok(())
    } else {
        Err(invalid_error(
            "flat ODC attribute is not valid for the selected element",
        ))
    }
}

fn exact_attribute(namespace: &ResolveResult<'_>, local: &[u8]) -> Option<ExactAttribute> {
    if resolved(namespace, CHART) {
        match local {
            b"class" => Some(ExactAttribute::Class),
            b"style-name" => Some(ExactAttribute::StyleName),
            b"values-cell-range-address" => Some(ExactAttribute::ValuesCellRangeAddress),
            b"label-cell-address" => Some(ExactAttribute::LabelCellAddress),
            b"attached-axis" => Some(ExactAttribute::AttachedAxis),
            _ => None,
        }
    } else if resolved(namespace, SVG) {
        match local {
            b"width" => Some(ExactAttribute::Width),
            b"height" => Some(ExactAttribute::Height),
            b"x" => Some(ExactAttribute::X),
            b"y" => Some(ExactAttribute::Y),
            _ => None,
        }
    } else if resolved(namespace, TABLE) && local == b"cell-range-address" {
        Some(ExactAttribute::CellRangeAddress)
    } else {
        None
    }
}

fn attribute_name(attribute: ExactAttribute) -> (&'static [u8], &'static str) {
    match attribute {
        ExactAttribute::Class => (CHART, "class"),
        ExactAttribute::StyleName => (CHART, "style-name"),
        ExactAttribute::Width => (SVG, "width"),
        ExactAttribute::Height => (SVG, "height"),
        ExactAttribute::X => (SVG, "x"),
        ExactAttribute::Y => (SVG, "y"),
        ExactAttribute::CellRangeAddress => (TABLE, "cell-range-address"),
        ExactAttribute::ValuesCellRangeAddress => (CHART, "values-cell-range-address"),
        ExactAttribute::LabelCellAddress => (CHART, "label-cell-address"),
        ExactAttribute::AttachedAxis => (CHART, "attached-axis"),
    }
}

fn parse(bytes: Vec<u8>, root_kind: RootKind, limits: crate::Limits) -> Result<State> {
    if bytes.len() > limits.max_content_bytes() {
        return Err(invalid_error(
            "flat ODC exceeds the caller-selected content limit",
        ));
    }
    if root_kind == RootKind::Flat
        && litchi_odf_common::detect::flat(&bytes) != Some(FileFormat::Odc)
    {
        return Err(invalid_error("input is not a flat ODC document"));
    }
    let xml =
        std::str::from_utf8(&bytes).map_err(|_utf8_error| invalid_error("ODC XML is not UTF-8"))?;
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_depth = None;
    let mut body_seen = false;
    let mut office_chart_depth = None;
    let mut office_chart_seen = false;
    let mut chart_depth = None;
    let mut chart_seen = false;
    let mut plot_depth = None;
    let mut plot_seen = false;
    let mut root_start_name = None;
    let mut root_end_name = None;
    let mut axes = Vec::new();
    let mut chart_tag = None;
    let mut plot_tag = None;
    let mut series_tags = Vec::new();

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_overflow| invalid_error("ODC XML event offset exceeds this platform"))?;
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid_error(format!("invalid ODC XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event_end = usize::try_from(reader.buffer_position())
            .map_err(|_overflow| invalid_error("ODC XML event offset exceeds this platform"))?;
        match event {
            Event::Start(element) => {
                let event_depth = checked_depth(depth, limits.max_depth())?;
                let local = element.local_name();
                if event_depth == 1 {
                    let expected = match root_kind {
                        RootKind::Flat => b"document".as_slice(),
                        RootKind::Content => b"document-content".as_slice(),
                    };
                    if root_seen || namespace != NamespaceKind::Office || local.as_ref() != expected
                    {
                        return Err(invalid_error(
                            "ODC requires one expected office document root",
                        ));
                    }
                    root_seen = true;
                    root_start_name = Some(event_name_span(
                        &bytes,
                        event_start..event_end,
                        element.name().as_ref(),
                    )?);
                } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
                    if event_depth != 2 || body_seen {
                        return Err(invalid_error("office:body is misplaced or duplicated"));
                    }
                    body_seen = true;
                    body_depth = Some(event_depth);
                } else if namespace == NamespaceKind::Office && local.as_ref() == b"chart" {
                    if body_depth != Some(event_depth - 1) || office_chart_seen {
                        return Err(invalid_error("office:chart is misplaced or duplicated"));
                    }
                    office_chart_seen = true;
                    office_chart_depth = Some(event_depth);
                } else if namespace == NamespaceKind::Chart && local.as_ref() == b"chart" {
                    if office_chart_depth != Some(event_depth - 1) || chart_seen {
                        return Err(invalid_error("chart:chart is misplaced or duplicated"));
                    }
                    chart_seen = true;
                    chart_depth = Some(event_depth);
                    chart_tag = Some(record_editable_tag(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        ExactTarget::Chart,
                        limits,
                    )?);
                } else if namespace == NamespaceKind::Chart && local.as_ref() == b"plot-area" {
                    if chart_depth != Some(event_depth - 1) || plot_seen {
                        return Err(invalid_error("chart:plot-area is misplaced or duplicated"));
                    }
                    plot_seen = true;
                    plot_depth = Some(event_depth);
                    plot_tag = Some(record_editable_tag(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        ExactTarget::PlotArea,
                        limits,
                    )?);
                } else if namespace == NamespaceKind::Chart
                    && local.as_ref() == b"axis"
                    && plot_depth == Some(event_depth - 1)
                {
                    push_axis(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        &mut axes,
                        limits,
                    )?;
                } else if namespace == NamespaceKind::Chart
                    && local.as_ref() == b"series"
                    && plot_depth == Some(event_depth - 1)
                {
                    let index = series_tags.len();
                    series_tags.push(record_editable_tag(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        ExactTarget::Series(index),
                        limits,
                    )?);
                }
                depth = event_depth;
            },
            Event::Empty(element) => {
                let event_depth = checked_depth(depth, limits.max_depth())?;
                let local = element.local_name();
                if event_depth == 1
                    || (namespace == NamespaceKind::Office
                        && matches!(local.as_ref(), b"body" | b"chart"))
                    || (namespace == NamespaceKind::Chart && local.as_ref() == b"chart")
                {
                    return Err(invalid_error("flat ODC required structure cannot be empty"));
                }
                if namespace == NamespaceKind::Chart && local.as_ref() == b"plot-area" {
                    if chart_depth != Some(event_depth - 1) || plot_seen {
                        return Err(invalid_error("chart:plot-area is misplaced or duplicated"));
                    }
                    plot_seen = true;
                    plot_tag = Some(record_editable_tag(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        ExactTarget::PlotArea,
                        limits,
                    )?);
                }
                if namespace == NamespaceKind::Chart
                    && local.as_ref() == b"axis"
                    && plot_depth == Some(event_depth - 1)
                {
                    push_axis(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        &mut axes,
                        limits,
                    )?;
                }
                if namespace == NamespaceKind::Chart
                    && local.as_ref() == b"series"
                    && plot_depth == Some(event_depth - 1)
                {
                    let index = series_tags.len();
                    series_tags.push(record_editable_tag(
                        &reader,
                        &element,
                        event_start..event_end,
                        &bytes,
                        ExactTarget::Series(index),
                        limits,
                    )?);
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid_error("flat ODC XML depth underflow"));
                }
                let local = element.local_name();
                if depth == 1 {
                    root_end_name = Some(event_name_span(
                        &bytes,
                        event_start..event_end,
                        element.name().as_ref(),
                    )?);
                    root_closed = true;
                }
                if plot_depth == Some(depth)
                    && namespace == NamespaceKind::Chart
                    && local.as_ref() == b"plot-area"
                {
                    plot_depth = None;
                }
                if chart_depth == Some(depth)
                    && namespace == NamespaceKind::Chart
                    && local.as_ref() == b"chart"
                {
                    chart_depth = None;
                }
                if office_chart_depth == Some(depth)
                    && namespace == NamespaceKind::Office
                    && local.as_ref() == b"chart"
                {
                    office_chart_depth = None;
                }
                if body_depth == Some(depth)
                    && namespace == NamespaceKind::Office
                    && local.as_ref() == b"body"
                {
                    body_depth = None;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid_error("DOCTYPE is not allowed in ODC")),
            Event::GeneralRef(reference)
                if !matches!(
                    reference.as_ref(),
                    b"amp" | b"lt" | b"gt" | b"apos" | b"quot"
                ) =>
            {
                return Err(invalid_error("flat ODC contains an unsupported entity"));
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::GeneralRef(_) => {},
        }
    }
    if depth != 0
        || !root_seen
        || !root_closed
        || !body_seen
        || !office_chart_seen
        || !chart_seen
        || body_depth.is_some()
        || office_chart_depth.is_some()
        || chart_depth.is_some()
        || plot_depth.is_some()
    {
        return Err(invalid_error("flat ODC structure is incomplete"));
    }

    let normalized_bytes = if root_kind == RootKind::Flat {
        let mut candidate = bytes.clone();
        let start =
            root_start_name.ok_or_else(|| invalid_error("flat ODC root start is missing"))?;
        let end = root_end_name.ok_or_else(|| invalid_error("flat ODC root end is missing"))?;
        let replacement = root_content_name(&bytes[start.clone()])?;
        for span in [end, start] {
            candidate.splice(span, replacement.iter().copied());
        }
        candidate
    } else {
        bytes.clone()
    };
    let normalized_xml = std::str::from_utf8(&normalized_bytes)
        .map_err(|_utf8_error| invalid_error("normalized ODC is not UTF-8"))?;
    let chart = chart::read(normalized_xml)?;
    let _ = chart.chart_class()?;
    crate::codec::validate_tree(&chart, limits)?;
    if let Some(plot_area) = chart.plot_area() {
        for axis in plot_area.axes() {
            let _ = axis.dimension()?;
        }
    }
    Ok(State {
        bytes,
        chart,
        axes,
        chart_tag: chart_tag.ok_or_else(|| invalid_error("flat ODC chart tag is missing"))?,
        plot_tag: plot_tag.ok_or_else(|| invalid_error("flat ODC plot-area tag is missing"))?,
        series_tags,
        root_kind,
        limits,
    })
}

fn record_editable_tag(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: Range<usize>,
    bytes: &[u8],
    target: ExactTarget,
    limits: crate::Limits,
) -> Result<EditableTag> {
    let mut attributes = Vec::new();
    for attribute_result in element.attributes().with_checks(true) {
        let source_attribute = attribute_result
            .map_err(|error| invalid_error(format!("invalid flat ODC attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(source_attribute.key);
        let Some(attribute) = exact_attribute(&namespace, local.as_ref()) else {
            continue;
        };
        if validate_exact_pair(target, attribute).is_err() {
            continue;
        }
        if attributes
            .iter()
            .any(|record: &AttributeRecord| record.attribute == attribute)
        {
            return Err(invalid_error(
                "flat ODC element has a duplicate controlled attribute",
            ));
        }
        let value = source_attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| invalid_error(format!("invalid flat ODC attribute value: {error}")))?
            .into_owned();
        if value.len() > MAX_NAME_BYTES || value.len() > limits.max_scalar_bytes() {
            return Err(invalid_error("flat ODC attribute exceeds its byte limit"));
        }
        let (value_span, attribute_span) =
            attribute_spans(&bytes[tag.clone()], source_attribute.key.as_ref())?;
        attributes.push(AttributeRecord {
            attribute,
            value,
            value_span: tag.start + value_span.start..tag.start + value_span.end,
            attribute_span: tag.start + attribute_span.start..tag.start + attribute_span.end,
        });
    }
    let prefix = element
        .name()
        .as_ref()
        .split(|byte| *byte == b':')
        .next()
        .filter(|_| element.name().as_ref().contains(&b':'))
        .map(|value| String::from_utf8_lossy(value).into_owned());
    Ok(EditableTag {
        tag,
        prefix,
        attributes,
    })
}

fn push_axis(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    tag: Range<usize>,
    bytes: &[u8],
    axes: &mut Vec<AxisRecord>,
    limits: crate::Limits,
) -> Result<()> {
    if axes.len() >= limits.max_axes() {
        return Err(invalid_error(
            "flat ODC axis count exceeds the caller-selected limit",
        ));
    }
    let mut name = None;
    let mut key = None;
    let mut style_name = None;
    let mut style_key = None;
    for attribute_result in element.attributes().with_checks(true) {
        let attribute = attribute_result
            .map_err(|error| invalid_error(format!("invalid flat ODC attribute: {error}")))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved(&namespace, CHART) && local.as_ref() == b"name" {
            if key.is_some() {
                return Err(invalid_error("flat ODC axis has duplicate chart:name"));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid flat ODC axis name: {error}")))?
                .into_owned();
            if value.len() > MAX_NAME_BYTES || value.len() > limits.max_scalar_bytes() {
                return Err(invalid_error("flat ODC axis name exceeds its byte limit"));
            }
            name = Some(value);
            key = Some(attribute.key.as_ref().to_vec());
        } else if resolved(&namespace, CHART) && local.as_ref() == b"style-name" {
            if style_key.is_some() {
                return Err(invalid_error(
                    "flat ODC axis has duplicate chart:style-name",
                ));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid_error(format!("invalid flat ODC axis style: {error}")))?
                .into_owned();
            if value.len() > MAX_NAME_BYTES || value.len() > limits.max_scalar_bytes() {
                return Err(invalid_error("flat ODC axis style exceeds its byte limit"));
            }
            style_name = Some(value);
            style_key = Some(attribute.key.as_ref().to_vec());
        }
    }
    let prefix = element
        .name()
        .as_ref()
        .split(|byte| *byte == b':')
        .next()
        .filter(|_| element.name().as_ref().contains(&b':'))
        .map(|value| String::from_utf8_lossy(value).into_owned());
    let (name_value, name_attribute) = if let Some(attribute_key) = key {
        let (value, attribute) = attribute_spans(&bytes[tag.clone()], &attribute_key)?;
        (
            Some(tag.start + value.start..tag.start + value.end),
            Some(tag.start + attribute.start..tag.start + attribute.end),
        )
    } else {
        (None, None)
    };
    let (style_value, style_attribute) = if let Some(attribute_key) = style_key {
        let (value, attribute) = attribute_spans(&bytes[tag.clone()], &attribute_key)?;
        (
            Some(tag.start + value.start..tag.start + value.end),
            Some(tag.start + attribute.start..tag.start + attribute.end),
        )
    } else {
        (None, None)
    };
    axes.push(AxisRecord {
        name,
        style_name,
        tag,
        name_value,
        name_attribute,
        style_value,
        style_attribute,
        prefix,
    });
    Ok(())
}

fn stage_axis_attribute(
    source: &FlatChart,
    axis: &AxisRecord,
    local: &str,
    after: Option<&String>,
    value: Option<&Range<usize>>,
    attribute: Option<&Range<usize>>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    match (after, value, attribute) {
        (Some(new_value), Some(value_span), _) => replacements.push(Replacement {
            range: value_span.clone(),
            value: escape_attribute(new_value).into_bytes(),
        }),
        (None, _, Some(attribute_span)) => replacements.push(Replacement {
            range: attribute_span.clone(),
            value: Vec::new(),
        }),
        (Some(new_value), None, None) => {
            let prefix = axis.prefix.as_deref().ok_or_else(|| {
                Error::Unsupported(format!(
                    "cannot add chart:{local} without a lossless chart namespace prefix"
                ))
            })?;
            let raw = &source.as_bytes()[axis.tag.clone()];
            let relative = insertion_offset(raw)?;
            replacements.push(Replacement {
                range: axis.tag.start + relative..axis.tag.start + relative,
                value: format!(" {prefix}:{local}=\"{}\"", escape_attribute(new_value))
                    .into_bytes(),
            });
        },
        (None, None, None) => {},
        (None, Some(_), None) | (Some(_), None, Some(_)) => {
            return Err(invalid_error("flat ODC axis attribute span is incomplete"));
        },
    }
    Ok(())
}

fn stage_exact_attribute(
    source: &FlatChart,
    tag: &EditableTag,
    attribute: ExactAttribute,
    after: Option<&String>,
    replacements: &mut Vec<Replacement>,
) -> Result<()> {
    let existing_record = tag.record(attribute);
    match (after, existing_record) {
        (Some(new_value), Some(attribute_record)) => replacements.push(Replacement {
            range: attribute_record.value_span.clone(),
            value: escape_attribute(new_value).into_bytes(),
        }),
        (None, Some(attribute_record)) => replacements.push(Replacement {
            range: attribute_record.attribute_span.clone(),
            value: Vec::new(),
        }),
        (Some(new_value), None) => {
            let (namespace, local) = attribute_name(attribute);
            if namespace != CHART {
                return Err(Error::Unsupported(format!(
                    "cannot add {local} without a lossless namespace-prefix proof"
                )));
            }
            let prefix = tag.prefix.as_deref().ok_or_else(|| {
                Error::Unsupported(format!(
                    "cannot add chart:{local} without a lossless chart namespace prefix"
                ))
            })?;
            let relative = insertion_offset(&source.as_bytes()[tag.tag.clone()])?;
            replacements.push(Replacement {
                range: tag.tag.start + relative..tag.tag.start + relative,
                value: format!(" {prefix}:{local}=\"{}\"", escape_attribute(new_value))
                    .into_bytes(),
            });
        },
        (None, None) => {},
    }
    Ok(())
}

fn attribute_spans(tag: &[u8], wanted: &[u8]) -> Result<(Range<usize>, Range<usize>)> {
    let mut cursor = 1usize;
    while cursor < tag.len() && !tag[cursor].is_ascii_whitespace() && tag[cursor] != b'>' {
        cursor += 1;
    }
    while cursor < tag.len() {
        let whitespace = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= tag.len() || matches!(tag[cursor], b'/' | b'>') {
            break;
        }
        let name_start = cursor;
        while cursor < tag.len()
            && !tag[cursor].is_ascii_whitespace()
            && !matches!(tag[cursor], b'=' | b'/' | b'>')
        {
            cursor += 1;
        }
        let name_end = cursor;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if tag.get(cursor) != Some(&b'=') {
            return Err(invalid_error("flat ODC attribute is missing '='"));
        }
        cursor += 1;
        while cursor < tag.len() && tag[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *tag
            .get(cursor)
            .filter(|quote| matches!(quote, b'\'' | b'\"'))
            .ok_or_else(|| invalid_error("flat ODC attribute is not quoted"))?;
        cursor += 1;
        let value_start = cursor;
        while cursor < tag.len() && tag[cursor] != quote {
            cursor += 1;
        }
        let value_end = cursor;
        cursor = cursor
            .checked_add(1)
            .ok_or_else(|| invalid_error("flat ODC attribute offset overflow"))?;
        if &tag[name_start..name_end] == wanted {
            return Ok((value_start..value_end, whitespace..cursor));
        }
    }
    Err(invalid_error("flat ODC axis name span was not found"))
}

fn insertion_offset(tag: &[u8]) -> Result<usize> {
    let close = tag
        .iter()
        .rposition(|byte| *byte == b'>')
        .ok_or_else(|| invalid_error("flat ODC axis start tag is incomplete"))?;
    Ok(if close > 0 && tag[close - 1] == b'/' {
        close - 1
    } else {
        close
    })
}

fn event_name_span(bytes: &[u8], event: Range<usize>, name: &[u8]) -> Result<Range<usize>> {
    let raw = bytes
        .get(event.clone())
        .ok_or_else(|| invalid_error("flat ODC event span is invalid"))?;
    let relative = raw
        .windows(name.len())
        .position(|candidate| candidate == name)
        .ok_or_else(|| invalid_error("flat ODC event name span is missing"))?;
    Ok(event.start + relative..event.start + relative + name.len())
}

fn root_content_name(name: &[u8]) -> Result<Vec<u8>> {
    let qualified_name = std::str::from_utf8(name)
        .map_err(|_utf8_error| invalid_error("flat ODC root name is invalid"))?;
    let prefix = qualified_name
        .strip_suffix("document")
        .ok_or_else(|| invalid_error("flat ODC root name is invalid"))?;
    Ok(format!("{prefix}document-content").into_bytes())
}

fn checked_depth(depth: usize, max_depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid_error("flat ODC XML depth overflow"))?;
    if next_depth > max_depth {
        return Err(invalid_error(
            "flat ODC XML depth exceeds the caller-selected limit",
        ));
    }
    Ok(next_depth)
}

fn resolved(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == CHART => NamespaceKind::Chart,
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
