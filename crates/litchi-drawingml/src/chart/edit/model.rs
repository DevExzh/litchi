//! Small typed edit vocabulary for chart subtrees.

/// A shared data-label switch in a series data-label container.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DataLabelFlag {
    /// Show the legend key beside the data label.
    ShowLegendKey,
    /// Show the point value.
    ShowValue,
    /// Show the category name.
    ShowCategoryName,
    /// Show the series name.
    ShowSeriesName,
    /// Show the percentage.
    ShowPercent,
    /// Show the bubble size.
    ShowBubbleSize,
    /// Show leader lines.
    ShowLeaderLines,
    /// Delete the data labels.
    Deleted,
}

impl DataLabelFlag {
    pub(crate) const fn element(self) -> &'static str {
        match self {
            Self::ShowLegendKey => "showLegendKey",
            Self::ShowValue => "showVal",
            Self::ShowCategoryName => "showCatName",
            Self::ShowSeriesName => "showSerName",
            Self::ShowPercent => "showPercent",
            Self::ShowBubbleSize => "showBubbleSize",
            Self::ShowLeaderLines => "showLeaderLines",
            Self::Deleted => "delete",
        }
    }
}
