//! Internal writer state for chart XML capabilities.

#[derive(Clone, Copy)]
pub(super) struct SeriesFeatures {
    pub(super) point_and_label_overrides: bool,
    pub(super) error_bars: bool,
    pub(super) trendlines: bool,
    pub(super) explosion: bool,
    pub(super) invert_if_negative: bool,
    pub(super) picture_options: bool,
    pub(super) bar_shape: bool,
    pub(super) marker: bool,
    pub(super) smooth: bool,
}

impl SeriesFeatures {
    pub(super) const BASIC: Self = Self {
        point_and_label_overrides: true,
        error_bars: true,
        trendlines: true,
        explosion: false,
        invert_if_negative: false,
        picture_options: false,
        bar_shape: false,
        marker: false,
        smooth: false,
    };
    pub(super) const AREA: Self = Self {
        picture_options: true,
        ..Self::BASIC
    };
    pub(super) const BAR: Self = Self {
        invert_if_negative: true,
        picture_options: true,
        bar_shape: true,
        ..Self::BASIC
    };
    pub(super) const LINE: Self = Self {
        marker: true,
        smooth: true,
        ..Self::BASIC
    };
    pub(super) const BUBBLE: Self = Self {
        invert_if_negative: true,
        ..Self::BASIC
    };
    pub(super) const LINE_3D: Self = Self {
        marker: true,
        ..Self::BASIC
    };
    pub(super) const PIE: Self = Self {
        explosion: true,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
    pub(super) const RADAR: Self = Self {
        marker: true,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
    pub(super) const SURFACE: Self = Self {
        point_and_label_overrides: false,
        error_bars: false,
        trendlines: false,
        ..Self::BASIC
    };
}
