use super::{
    ConditionalColorScale, ConditionalColorScaleEntry, ConditionalDataBar, ConditionalFormatRule,
    ConditionalIconSet, Link, SparklineComplexColor, SparklineGroup,
};

/// A `text:a` hyperlink whose text content is still being collected.
pub(super) struct PendingHyperlink {
    /// The hyperlink parsed from the element's attributes.
    pub(super) link: Link,
    /// Byte offset into the cell text where the link text begins.
    pub(super) text_start: usize,
    /// The `text_element_depth` value assigned to the `text:a` element.
    pub(super) depth: usize,
}

/// A `calcext:conditional-format` element whose rules are still being read.
pub(super) struct PendingConditionalFormat {
    /// Target ranges parsed from `calcext:target-range-address`.
    pub(super) target_range_addresses: Vec<String>,
    /// Inert rules collected so far, in document order.
    pub(super) rules: Vec<ConditionalFormatRule>,
    /// The `element_depth` value assigned to the element.
    pub(super) depth: usize,
}

/// A `calcext:color-scale`, `calcext:data-bar`, or `calcext:icon-set` rule
/// whose threshold entries are still being read.
pub(super) enum PendingCalcextRule {
    ColorScale {
        entries: Vec<ConditionalColorScaleEntry>,
        depth: usize,
    },
    DataBar {
        data_bar: ConditionalDataBar,
        depth: usize,
    },
    IconSet {
        icon_set: ConditionalIconSet,
        depth: usize,
    },
}

impl PendingCalcextRule {
    pub(super) fn depth(&self) -> usize {
        match self {
            Self::ColorScale { depth, .. }
            | Self::DataBar { depth, .. }
            | Self::IconSet { depth, .. } => *depth,
        }
    }

    pub(super) fn element_name(&self) -> &'static str {
        match self {
            Self::ColorScale { .. } => "color-scale",
            Self::DataBar { .. } => "data-bar",
            Self::IconSet { .. } => "icon-set",
        }
    }

    pub(super) fn finish(self) -> ConditionalFormatRule {
        match self {
            Self::ColorScale { entries, .. } => ConditionalColorScale::new(entries).into(),
            Self::DataBar { data_bar, .. } => data_bar.into(),
            Self::IconSet { icon_set, .. } => icon_set.into(),
        }
    }
}

/// A `calcext:sparkline-group` element whose sparklines are still being read.
pub(super) struct PendingSparklineGroup {
    /// The group parsed from the element's attributes (with no sparklines yet).
    pub(super) group: SparklineGroup,
    /// The `element_depth` value assigned to the element.
    pub(super) depth: usize,
}

/// A `calcext:sparkline-*-complex-color` element whose `loext:transformation`
/// children are still being read.
pub(super) struct PendingSparklineComplexColor {
    /// The slot element name (one of `COMPLEX_COLOR_SLOTS`).
    pub(super) slot: &'static str,
    /// The color parsed from the element's attributes (no transformations yet).
    pub(super) color: SparklineComplexColor,
    /// The `element_depth` value assigned to the element.
    pub(super) depth: usize,
}

#[derive(Clone, Copy)]
pub(super) enum SheetTextField {
    Title,
    Description,
}

impl SheetTextField {
    pub(super) fn local_name(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Description => "desc",
        }
    }
}
