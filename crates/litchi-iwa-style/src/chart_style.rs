//! The chart-level and axis-level style properties of a chart (records
//! 0375, 0378, 0414): the inspector values stored in the generated
//! extension (field 10,000) of a chart's `TSCH.ChartStyleArchive`,
//! `TSCH.ChartNonStyleArchive`, `TSCH.ChartAxisStyleArchive` and
//! `TSCH.ChartAxisNonStyleArchive`, plus the legend frame of the chart
//! drawable itself.
//!
//! The property values ([`chart::Property`], [`axis::Property`]) are the
//! shared vocabulary of the three format families; the mapping of each
//! value to its stored fields is the host editors' `set_*_chart_*` routes
//! (`charts::{pie_start_angle, donut_inner_radius, radar_start_angle,
//! series_gap_3d, scene_3d, depth_3d, rounded_corners, legend_frame,
//! options, border, hidden_data, bar_shape_3d, radar_grid_shape, gaps,
//! axis_style}`) on the bounded chart style codec.
//!
//! Two layers, as for the series styles:
//!
//! - payload functions ([`chart::value`], [`chart::rewrite`],
//!   [`axis::value`], [`axis::rewrite`]) for an owner that resolves the
//!   style object and publishes its own in-place rewrite;
//! - workspace functions ([`read_chart`], [`write_chart`], [`read_axis`],
//!   [`write_axis`]) that resolve the chart's style object over a
//!   [`Workspace`], refuse a style another chart shares (the host's
//!   `ensure_exclusive`), rewrite it in place under a residual [`Limits`]
//!   profile and report the source-derived mutation scope.

use litchi_iwa_common::chart::axis::Axis;
use litchi_iwa_common::chart::axis::style::Visibility;
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits as CodecLimits, Properties, StyleReferences,
};

use litchi_iwa_common::shape::shadow::Shadow as StoredShadow;

use crate::graph::{Error, Workspace};
use crate::series::{
    Limits, MutationScope, Refusal, Report, SymbolBudget, SymbolError, SymbolResult,
    bounded_chart_payload, bounded_payload, bounded_references, charged_patch,
};

/// `TSCH.ChartDrawableArchive` message type.
const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ChartArchive.legend_frame`, in the chart drawable's extension.
const LEGEND_FRAME_FIELD: u32 = 3;

fn codec_error(error: codec::Error) -> SymbolError {
    SymbolError::Graph(Error::Codec(error))
}

/// A stored value outside its semantic range.
fn out_of_range<E>(_error: E) -> SymbolError {
    SymbolError::Invalid("a stored chart style value is out of range")
}

fn decode(payload: &[u8], archive: Archive) -> SymbolResult<Properties<'_>> {
    codec::decode(payload, archive, CodecLimits::DEFAULT).map_err(codec_error)
}

fn apply_writes(
    payload: &[u8],
    archive: Archive,
    writes: &[codec::Write<'_>],
    empty: codec::EmptyExtension,
) -> SymbolResult<Vec<u8>> {
    Ok(
        codec::rewrite(payload, archive, writes, empty, CodecLimits::DEFAULT)
            .map_err(codec_error)?
            .into_bytes()
            .into_owned(),
    )
}

mod lighting_package;
mod lighting_presets;

/// The chart-level properties.
pub mod chart {
    use litchi_iwa_common::chart::kind::Kind as ChartKind;
    use litchi_iwa_common::chart::{corner, donut, gaps, legend, pie, radar};
    use litchi_iwa_common::chart3d::{self, Scale};
    use litchi_iwa_common::color::{RgbColorSpace, Rgba};
    use litchi_iwa_common::shape::fill::{Angle, Fill, Gradient};
    use litchi_iwa_common::shape::shadow::Drop;
    use litchi_iwa_common::shape::stroke::{Pattern, Stroke, Width};
    use litchi_iwa_protos::chart_style_codec::{
        self as codec, Archive, Limits as CodecLimits, Properties, Write, field, property,
    };
    use litchi_iwa_protos::shape_style_codec::native as stored;

    use crate::drawing;

    /// The fill a chart style without a stored background fill draws (the
    /// host's `native_default_chart_background_fill`: a linear gradient from
    /// 67% gray to white at 0 degrees).
    #[must_use]
    pub fn default_background() -> Fill {
        let gray = 2.0 / 3.0;
        Fill::Gradient(Gradient::linear(
            Rgba::new(gray, gray, gray, 1.0, RgbColorSpace::Srgb).expect("opaque gray"),
            Rgba::new(1.0, 1.0, 1.0, 1.0, RgbColorSpace::Srgb).expect("opaque white"),
            Angle::from_degrees(0.0).expect("zero angle"),
        ))
    }

    /// The chart border a chart style without a stored border draws (the
    /// host's `native_default_chart_border_stroke`: black, one point, solid).
    #[must_use]
    pub fn default_border() -> Stroke {
        Stroke::new(Rgba::black(), Width::ONE, Pattern::Solid)
    }

    use super::{apply_writes, codec_error, out_of_range};
    use crate::series::{Refusal, SymbolError, SymbolResult};

    fn decode_chart(payload: &[u8], archive: Archive) -> SymbolResult<Properties<'_>> {
        codec::decode(payload, archive, CodecLimits::DEFAULT).map_err(codec_error)
    }

    /// One chart-level style property with its value.
    #[derive(Debug, Clone, PartialEq)]
    #[non_exhaustive]
    pub enum Property {
        /// Wedges rotation of a pie or donut chart.
        PieStartAngle(pie::StartAngle),
        /// Segment inner radius of a donut chart.
        DonutInnerRadius(donut::InnerRadius),
        /// Rotation of a radar chart.
        RadarStartAngle(radar::StartAngle),
        /// Gap between series of a 3D line or area chart.
        SeriesGap3d(chart3d::SeriesGap),
        /// Scene rotation of a 3D chart.
        Rotation3d(chart3d::Rotation),
        /// Depth of a 3D chart.
        Depth3d(chart3d::Depth),
        /// Rounded corners of bars and columns.
        RoundedCorners(corner::Rounding),
        /// Placement of the legend (automatic or an explicit frame).
        LegendFrame(legend::Frame),
        /// Whether the legend shows.
        LegendVisible(bool),
        /// Whether the chart area shows its border.
        BorderVisible(bool),
        /// Whether data in hidden table rows and columns is plotted.
        IncludesHiddenData(bool),
        /// The bar shape of a 3D bar or column chart.
        BarShape3d(chart3d::BarShape),
        /// The radius gridlines of a radar chart.
        RadarGridShape(RadarGridShape),
        /// The gaps between bars or columns and between their sets.
        GapSpacing(gaps::Spacing),
        /// The chart-area border stroke (`None`: no border line).
        BorderStroke(Option<Stroke>),
        /// The fill behind the plot area. Image fills move data references
        /// and media and are refused here.
        BackgroundFill(Fill),
        /// The lighting preset of a 3D chart.
        Lighting3d(Lighting3d),
        /// The drop shadow of the chart's series, drawn per series or
        /// grouped. Stored in every series style and the chart style.
        Shadow(Shadow),
        /// The chart-wide appearance of a radar chart's series, stored in
        /// every series style.
        RadarSeriesStyle(RadarSeriesStyle),
    }

    /// The Radar inspector's chart-wide series "Style" control (the host's
    /// `ChartRadarSeriesStyle`): one coordinated fill and stroke
    /// representation applied to every series.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    #[non_exhaustive]
    pub enum RadarSeriesStyle {
        /// Translucent series-color fills without outlines.
        Fill,
        /// Series-color outlines without an area fill.
        Stroke,
        /// Series-color outlines with a derived translucent area fill (the
        /// native default).
        #[default]
        FillAndStroke,
    }

    /// The shadow of a chart's series (the host's `ChartShadow`).
    #[derive(Debug, Clone, Copy, PartialEq)]
    #[non_exhaustive]
    pub enum Shadow {
        /// No shadow.
        None,
        /// A drop shadow drawn behind each series.
        IndividualSeries(Drop),
        /// A drop shadow drawn behind the series grouped as one layer.
        Grouped(Drop),
    }

    impl Shadow {
        /// The shadow of a chart without stored shadows: a black drop
        /// shadow per series (10 pt blur, 6 pt offset, 75% opacity, 45°).
        pub const NATIVE_DEFAULT: Self =
            Self::IndividualSeries(crate::series::style::NATIVE_DEFAULT_SHADOW);

        /// The drop shadow drawn, if any.
        #[must_use]
        pub const fn drop_shadow(self) -> Option<Drop> {
            match self {
                Self::None => None,
                Self::IndividualSeries(shadow) | Self::Grouped(shadow) => Some(shadow),
            }
        }

        /// Whether the series draw as one grouped layer.
        #[must_use]
        pub const fn groups_series(self) -> bool {
            matches!(self, Self::Grouped(_))
        }
    }

    impl Default for Shadow {
        fn default() -> Self {
            Self::NATIVE_DEFAULT
        }
    }

    /// One native choice from the Chart inspector's Lighting Style menu of
    /// a 3D chart (the host's `Chart3dLightingStyle`).
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    #[non_exhaustive]
    pub enum Lighting3d {
        /// No stored lighting package (the native default).
        #[default]
        Default,
        /// "Soft Light".
        SoftLight,
        /// "Soft Fill".
        SoftFill,
        /// "Medium Center".
        MediumCenter,
        /// "Medium Right".
        MediumRight,
        /// "Medium Left".
        MediumLeft,
        /// "Glossy".
        Glossy,
    }

    impl Lighting3d {
        const ALL: [Self; 7] = [
            Self::Default,
            Self::SoftLight,
            Self::SoftFill,
            Self::MediumCenter,
            Self::MediumRight,
            Self::MediumLeft,
            Self::Glossy,
        ];

        /// The label the native Chart inspector shows (the stored package
        /// name).
        #[must_use]
        pub const fn inspector_label(self) -> &'static str {
            match self {
                Self::Default => "Default",
                Self::SoftLight => "Soft Light",
                Self::SoftFill => "Soft Fill",
                Self::MediumCenter => "Medium Center",
                Self::MediumRight => "Medium Right",
                Self::MediumLeft => "Medium Left",
                Self::Glossy => "Glossy",
            }
        }

        /// The encoded package the host writes; `None` for the default,
        /// which removes the package.
        const fn package(self) -> Option<&'static [u8]> {
            use super::lighting_presets as preset;
            Some(match self {
                Self::Default => return None,
                Self::SoftLight => preset::SOFT_LIGHT,
                Self::SoftFill => preset::SOFT_FILL,
                Self::MediumCenter => preset::MEDIUM_CENTER,
                Self::MediumRight => preset::MEDIUM_RIGHT,
                Self::MediumLeft => preset::MEDIUM_LEFT,
                Self::Glossy => preset::GLOSSY,
            })
        }

        /// The preset a stored package names. The host refuses a package
        /// without a light, with an invalid light or with a name outside
        /// the menu.
        fn from_package(package: &[u8]) -> SymbolResult<Self> {
            let name = super::lighting_package::validated_name(package)?;
            Self::ALL
                .into_iter()
                .find(|preset| preset.inspector_label() == name)
                .ok_or(SymbolError::Invalid("an unsupported 3D lighting package"))
        }
    }

    /// The chart style extension field storing the lighting package of a
    /// chart of kind `chart` (the host's `LightingField`).
    const fn lighting_field(chart: ChartKind) -> Option<u32> {
        Some(match chart {
            ChartKind::Area3d | ChartKind::StackedArea3d => 2,
            ChartKind::Bar3d | ChartKind::StackedBar3d => 3,
            ChartKind::Column3d | ChartKind::StackedColumn3d => 4,
            ChartKind::Line3d => 6,
            ChartKind::Pie3d | ChartKind::Donut3d => 7,
            _ => return None,
        })
    }

    /// The radius gridlines of a radar chart.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
    #[non_exhaustive]
    pub enum RadarGridShape {
        /// Straight segments forming a polygon (the native default).
        #[default]
        Straight,
        /// Circular arcs.
        Curved,
    }

    /// The property a [`Property`] value sets, for reads and resets.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    #[non_exhaustive]
    pub enum PropertyKind {
        /// [`Property::PieStartAngle`].
        PieStartAngle,
        /// [`Property::DonutInnerRadius`].
        DonutInnerRadius,
        /// [`Property::RadarStartAngle`].
        RadarStartAngle,
        /// [`Property::SeriesGap3d`].
        SeriesGap3d,
        /// [`Property::Rotation3d`].
        Rotation3d,
        /// [`Property::Depth3d`].
        Depth3d,
        /// [`Property::RoundedCorners`].
        RoundedCorners,
        /// [`Property::LegendFrame`].
        LegendFrame,
        /// [`Property::LegendVisible`].
        LegendVisible,
        /// [`Property::BorderVisible`].
        BorderVisible,
        /// [`Property::IncludesHiddenData`].
        IncludesHiddenData,
        /// [`Property::BarShape3d`].
        BarShape3d,
        /// [`Property::RadarGridShape`].
        RadarGridShape,
        /// [`Property::GapSpacing`].
        GapSpacing,
        /// [`Property::BorderStroke`].
        BorderStroke,
        /// [`Property::BackgroundFill`].
        BackgroundFill,
        /// [`Property::Lighting3d`].
        Lighting3d,
        /// [`Property::Shadow`].
        Shadow,
        /// [`Property::RadarSeriesStyle`].
        RadarSeriesStyle,
    }

    impl Property {
        /// The property this value sets.
        #[must_use]
        pub const fn kind(&self) -> PropertyKind {
            match self {
                Self::PieStartAngle(_) => PropertyKind::PieStartAngle,
                Self::DonutInnerRadius(_) => PropertyKind::DonutInnerRadius,
                Self::RadarStartAngle(_) => PropertyKind::RadarStartAngle,
                Self::SeriesGap3d(_) => PropertyKind::SeriesGap3d,
                Self::Rotation3d(_) => PropertyKind::Rotation3d,
                Self::Depth3d(_) => PropertyKind::Depth3d,
                Self::RoundedCorners(_) => PropertyKind::RoundedCorners,
                Self::LegendFrame(_) => PropertyKind::LegendFrame,
                Self::LegendVisible(_) => PropertyKind::LegendVisible,
                Self::BorderVisible(_) => PropertyKind::BorderVisible,
                Self::IncludesHiddenData(_) => PropertyKind::IncludesHiddenData,
                Self::BarShape3d(_) => PropertyKind::BarShape3d,
                Self::RadarGridShape(_) => PropertyKind::RadarGridShape,
                Self::GapSpacing(_) => PropertyKind::GapSpacing,
                Self::BorderStroke(_) => PropertyKind::BorderStroke,
                Self::BackgroundFill(_) => PropertyKind::BackgroundFill,
                Self::Lighting3d(_) => PropertyKind::Lighting3d,
                Self::Shadow(_) => PropertyKind::Shadow,
                Self::RadarSeriesStyle(_) => PropertyKind::RadarSeriesStyle,
            }
        }
    }

    impl PropertyKind {
        /// The native default of this property (what a reset writes).
        #[must_use]
        pub fn native_default(self) -> Property {
            match self {
                Self::PieStartAngle => Property::PieStartAngle(pie::StartAngle::ZERO),
                Self::DonutInnerRadius => Property::DonutInnerRadius(donut::InnerRadius::DEFAULT),
                Self::RadarStartAngle => Property::RadarStartAngle(radar::StartAngle::ZERO),
                Self::SeriesGap3d => Property::SeriesGap3d(chart3d::SeriesGap::NATIVE_DEFAULT),
                Self::Rotation3d => Property::Rotation3d(chart3d::Rotation::DEFAULT),
                Self::Depth3d => Property::Depth3d(chart3d::Depth::DEFAULT),
                Self::RoundedCorners => Property::RoundedCorners(corner::Rounding::NONE),
                Self::LegendFrame => Property::LegendFrame(legend::Frame::Automatic),
                Self::LegendVisible => Property::LegendVisible(false),
                Self::BorderVisible => Property::BorderVisible(false),
                Self::IncludesHiddenData => Property::IncludesHiddenData(true),
                Self::BarShape3d => Property::BarShape3d(chart3d::BarShape::Rectangle),
                Self::RadarGridShape => Property::RadarGridShape(RadarGridShape::Straight),
                Self::GapSpacing => Property::GapSpacing(gaps::Spacing::DEFAULT),
                Self::BorderStroke => Property::BorderStroke(Some(default_border())),
                Self::BackgroundFill => Property::BackgroundFill(default_background()),
                Self::Lighting3d => Property::Lighting3d(Lighting3d::Default),
                Self::Shadow => Property::Shadow(Shadow::NATIVE_DEFAULT),
                Self::RadarSeriesStyle => {
                    Property::RadarSeriesStyle(RadarSeriesStyle::FillAndStroke)
                },
            }
        }

        /// The style archive storing the property; `None` for the legend
        /// frame, which the chart drawable stores. The radar series style
        /// is stored by the series styles.
        #[must_use]
        pub const fn archive(self) -> Option<Archive> {
            Some(match self {
                Self::LegendFrame => return None,
                Self::RadarSeriesStyle => Archive::SeriesStyle,
                Self::RoundedCorners
                | Self::BorderVisible
                | Self::RadarGridShape
                | Self::GapSpacing
                | Self::BorderStroke
                | Self::BackgroundFill
                | Self::Lighting3d
                | Self::Shadow => Archive::ChartStyle,
                _ => Archive::ChartNonStyle,
            })
        }

        /// Whether the property applies to a chart of kind `kind`.
        #[must_use]
        pub const fn applies_to(self, kind: ChartKind) -> bool {
            match self {
                Self::PieStartAngle => kind.supports_pie_start_angle(),
                Self::DonutInnerRadius => kind.supports_donut_inner_radius(),
                Self::RadarStartAngle => kind.supports_radar_start_angle(),
                Self::SeriesGap3d => kind.supports_3d_series_gap(),
                Self::Rotation3d => kind.supports_3d_scene(),
                Self::Depth3d => kind.supports_3d_depth(),
                Self::BarShape3d => kind.supports_3d_bar_shape(),
                Self::RadarGridShape => kind.supports_radar_grid_shape(),
                Self::RadarSeriesStyle => kind.supports_radar_series_style(),
                Self::Lighting3d => lighting_field(kind).is_some(),
                Self::RoundedCorners
                | Self::LegendFrame
                | Self::LegendVisible
                | Self::BorderVisible
                | Self::IncludesHiddenData
                | Self::GapSpacing
                | Self::BorderStroke
                | Self::BackgroundFill
                | Self::Shadow => true,
            }
        }
    }

    /// The extension fields of the chart-level scalar properties (the
    /// host's `charts::{options, border, hidden_data, bar_shape_3d,
    /// radar_grid_shape, gaps}` constants).
    mod scalar {
        /// `tschchartinfodefaultshowlegend` (chart non-style).
        pub(super) const LEGEND_VISIBLE: u32 = 20;
        /// `tschchartinfodefaultskiphiddendata` (chart non-style).
        pub(super) const SKIP_HIDDEN_DATA: u32 = 22;
        /// `tschchartinfodefault3dbarshape` (chart non-style).
        pub(super) const BAR_SHAPE_3D: u32 = 1;
        /// `tschchartinfodefaultshowborder` (chart style).
        pub(super) const BORDER_VISIBLE: u32 = 18;
        /// `tschchartinfodefaultradarradiusgridlinecurve` (chart style).
        pub(super) const RADAR_GRID_CURVE: u32 = 29;
        /// `tschchartinfodefaultgridbackgroundfill` (chart style).
        pub(super) const BACKGROUND_FILL: u32 = 14;
        /// `tschchartinfodefaultcombinelayers` (chart style).
        pub(super) const COMBINE_LAYERS: u32 = 13;
        /// `tschchartinfodefaultborderstroke` (chart style).
        pub(super) const BORDER_STROKE: u32 = 12;
        /// `tschchartinfodefaultinterbargap` and `…intersetgap` (chart
        /// style).
        pub(super) const GAP_BETWEEN_ITEMS: u32 = 16;
        pub(super) const GAP_BETWEEN_SETS: u32 = 17;
    }

    const fn scale3d(scale: Scale) -> property::Scale3d {
        match scale {
            Scale::Area => property::Scale3d::Area,
            Scale::Bar => property::Scale3d::Bar,
            Scale::Column => property::Scale3d::Column,
            Scale::Line => property::Scale3d::Line,
            _ => property::Scale3d::Pie,
        }
    }

    fn float(properties: &Properties<'_>, field: u32) -> SymbolResult<Option<f32>> {
        properties.float(field).map_err(codec_error)
    }

    fn flag(properties: &Properties<'_>, field: u32) -> SymbolResult<Option<bool>> {
        properties.bool(field).map_err(codec_error)
    }

    /// The stored (or native default) value of `kind` in the decoded style
    /// payload of its [`PropertyKind::archive`], for a chart of kind
    /// `chart`.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedKind`] for a 3D depth of a chart without a 3D
    /// scale; a stored value outside its range; the legend frame (read with
    /// [`legend_frame`]).
    pub fn value(
        properties: &Properties<'_>,
        kind: PropertyKind,
        chart: ChartKind,
    ) -> SymbolResult<Property> {
        use field::{chart_non_style as non_style, chart_style as style};
        Ok(match kind {
            PropertyKind::PieStartAngle => Property::PieStartAngle(
                float(properties, non_style::PIE_START_ANGLE)?
                    .map(pie::StartAngle::from_degrees)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or_default(),
            ),
            PropertyKind::DonutInnerRadius => Property::DonutInnerRadius(
                float(properties, non_style::DONUT_INNER_RADIUS)?
                    .map(donut::InnerRadius::from_native)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or_default(),
            ),
            PropertyKind::RadarStartAngle => Property::RadarStartAngle(
                float(properties, non_style::RADAR_START_ANGLE)?
                    .map(radar::StartAngle::from_degrees)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or_default(),
            ),
            PropertyKind::SeriesGap3d => Property::SeriesGap3d(
                float(properties, non_style::SERIES_GAP_3D)?
                    .map(chart3d::SeriesGap::from_native_ratio)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or_default(),
            ),
            PropertyKind::Rotation3d => Property::Rotation3d(
                properties
                    .vector(non_style::ROTATION_3D)
                    .map_err(codec_error)?
                    .map(|v| chart3d::Rotation::from_native([v.x, v.y, v.z, v.w]))
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or_default(),
            ),
            PropertyKind::Depth3d => {
                let scale =
                    Scale::for_kind(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
                Property::Depth3d(
                    properties
                        .vector(scale3d(scale).field())
                        .map_err(codec_error)?
                        .map(|v| chart3d::Depth::from_native(scale, [v.x, v.y, v.z, v.w]))
                        .transpose()
                        .map_err(out_of_range)?
                        .unwrap_or_default(),
                )
            },
            PropertyKind::RoundedCorners => Property::RoundedCorners(corner::Rounding::new(
                float(properties, style::ROUNDED_CORNER_RADIUS)?
                    .map(corner::Radius::from_native)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or(corner::Radius::ZERO),
                flag(properties, style::ROUNDED_CORNER_OUTER_END_ONLY)?.unwrap_or(false),
            )),
            PropertyKind::LegendVisible => {
                Property::LegendVisible(flag(properties, scalar::LEGEND_VISIBLE)?.unwrap_or(false))
            },
            PropertyKind::BorderVisible => {
                Property::BorderVisible(flag(properties, scalar::BORDER_VISIBLE)?.unwrap_or(false))
            },
            PropertyKind::IncludesHiddenData => Property::IncludesHiddenData(
                !flag(properties, scalar::SKIP_HIDDEN_DATA)?.unwrap_or(false),
            ),
            PropertyKind::BarShape3d => Property::BarShape3d(
                properties
                    .int32(scalar::BAR_SHAPE_3D)
                    .map_err(codec_error)?
                    .map_or(chart3d::BarShape::Rectangle, chart3d::BarShape::from_native),
            ),
            PropertyKind::RadarGridShape => Property::RadarGridShape(
                if flag(properties, scalar::RADAR_GRID_CURVE)? == Some(true) {
                    RadarGridShape::Curved
                } else {
                    RadarGridShape::Straight
                },
            ),
            PropertyKind::GapSpacing => {
                let gap = |field: u32, default: gaps::Percentage| -> SymbolResult<_> {
                    Ok(match float(properties, field)? {
                        Some(value) => gaps::Percentage::new(value).map_err(out_of_range)?,
                        None => default,
                    })
                };
                Property::GapSpacing(gaps::Spacing::new(
                    gap(
                        scalar::GAP_BETWEEN_ITEMS,
                        gaps::Spacing::DEFAULT.between_items(),
                    )?,
                    gap(
                        scalar::GAP_BETWEEN_SETS,
                        gaps::Spacing::DEFAULT.between_sets(),
                    )?,
                ))
            },
            PropertyKind::BorderStroke => Property::BorderStroke(
                match properties
                    .message(scalar::BORDER_STROKE)
                    .map_err(codec_error)?
                {
                    None => Some(default_border()),
                    Some(bytes) => drawing::stroke(
                        &stored::stroke(bytes, CodecLimits::DEFAULT).map_err(out_of_range)?,
                    )
                    .map_err(out_of_range)?,
                },
            ),
            PropertyKind::BackgroundFill => Property::BackgroundFill(
                match properties
                    .message(scalar::BACKGROUND_FILL)
                    .map_err(codec_error)?
                {
                    None => default_background(),
                    Some(bytes) => drawing::fill(
                        &stored::fill(bytes, CodecLimits::DEFAULT).map_err(out_of_range)?,
                    )
                    .map_err(out_of_range)?,
                },
            ),
            PropertyKind::Lighting3d => {
                let field =
                    lighting_field(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
                Property::Lighting3d(match properties.message(field).map_err(codec_error)? {
                    None => Lighting3d::Default,
                    Some(package) => Lighting3d::from_package(package)?,
                })
            },
            PropertyKind::Shadow => {
                return Err(SymbolError::Invalid(
                    "the chart shadow is stored by the series styles too",
                ));
            },
            PropertyKind::RadarSeriesStyle => {
                return Err(SymbolError::Invalid(
                    "the radar series style is stored by the series styles",
                ));
            },
            PropertyKind::LegendFrame => {
                return Err(SymbolError::Invalid(
                    "the legend frame is stored by the chart drawable",
                ));
            },
        })
    }

    /// The chart shadow from the decoded chart style payload and the
    /// uniform shadow of the chart's series styles (the host's
    /// `chart_shadow_from_native`).
    ///
    /// # Errors
    ///
    /// A disabled shadow over grouped layers, or a non-drop series shadow.
    pub fn shadow(
        properties: &Properties<'_>,
        series: litchi_iwa_common::shape::shadow::Shadow,
    ) -> SymbolResult<Shadow> {
        use litchi_iwa_common::shape::shadow::Shadow as Stored;
        let combined = flag(properties, scalar::COMBINE_LAYERS)?.unwrap_or(false);
        match series {
            Stored::Disabled if !combined => Ok(Shadow::None),
            Stored::Drop(drop) if combined => Ok(Shadow::Grouped(drop)),
            Stored::Drop(drop) => Ok(Shadow::IndividualSeries(drop)),
            Stored::Disabled => Err(SymbolError::Invalid(
                "a disabled chart shadow combines series layers",
            )),
            _ => Err(SymbolError::Invalid(
                "a chart series uses a non-drop shadow",
            )),
        }
    }

    /// The writes of a chart-level scalar property, as the host patches it.
    fn scalar_writes(value: Property) -> Option<Vec<Write<'static>>> {
        Some(match value {
            Property::LegendVisible(visible) => vec![Write::Bool {
                field: scalar::LEGEND_VISIBLE,
                value: Some(visible),
            }],
            Property::BorderVisible(visible) => vec![Write::BoolUnlessAbsentDefault {
                field: scalar::BORDER_VISIBLE,
                value: visible,
                default: false,
            }],
            Property::IncludesHiddenData(included) => vec![Write::BoolUnlessAbsentDefault {
                field: scalar::SKIP_HIDDEN_DATA,
                value: !included,
                default: false,
            }],
            Property::BarShape3d(shape) => vec![Write::Int32 {
                field: scalar::BAR_SHAPE_3D,
                value: (shape != chart3d::BarShape::Rectangle).then_some(shape.native_value()),
            }],
            Property::RadarGridShape(shape) => vec![Write::Bool {
                field: scalar::RADAR_GRID_CURVE,
                value: (shape == RadarGridShape::Curved).then_some(true),
            }],
            Property::GapSpacing(spacing) => {
                if spacing == gaps::Spacing::DEFAULT {
                    vec![
                        Write::Float {
                            field: scalar::GAP_BETWEEN_ITEMS,
                            value: None,
                        },
                        Write::Float {
                            field: scalar::GAP_BETWEEN_SETS,
                            value: None,
                        },
                    ]
                } else {
                    vec![
                        Write::FloatUnlessAbsentDefault {
                            field: scalar::GAP_BETWEEN_ITEMS,
                            value: spacing.between_items().percent(),
                            default: gaps::Spacing::DEFAULT.between_items().percent(),
                        },
                        Write::FloatUnlessAbsentDefault {
                            field: scalar::GAP_BETWEEN_SETS,
                            value: spacing.between_sets().percent(),
                            default: gaps::Spacing::DEFAULT.between_sets().percent(),
                        },
                    ]
                }
            },
            _ => return None,
        })
    }

    /// The codec's library write rule for a geometry `value`.
    fn rule(value: Property, chart: ChartKind) -> SymbolResult<property::PropertyWrite<'static>> {
        Ok(match value {
            Property::PieStartAngle(angle) => property::pie_start_angle(angle.degrees()),
            Property::DonutInnerRadius(radius) => {
                property::donut_inner_radius(radius.native_fraction())
            },
            Property::RadarStartAngle(angle) => property::radar_start_angle(angle.degrees()),
            Property::SeriesGap3d(gap) => property::series_gap_3d(gap.native_ratio()),
            Property::Rotation3d(rotation) => {
                property::rotation_3d(rotation.x_degrees(), rotation.y_degrees())
            },
            Property::Depth3d(depth) => {
                let scale =
                    Scale::for_kind(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
                property::depth_3d(
                    scale3d(scale),
                    depth.native_ratio(scale),
                    (depth != chart3d::Depth::DEFAULT).then(|| depth.native_vector(scale)),
                )
            },
            Property::RoundedCorners(rounding) => property::rounded_corners(
                rounding.radius().native_fraction(),
                rounding.outside_corners_only(),
            ),
            _ => return Err(SymbolError::Invalid("not a chart geometry property")),
        })
    }

    /// The payload of the style object of `value`'s
    /// [`PropertyKind::archive`] storing `value`, every other byte kept.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedKind`], an invalid payload, or the legend
    /// frame (written with [`rewrite_legend_frame`]).
    pub fn rewrite(payload: &[u8], value: &Property, chart: ChartKind) -> SymbolResult<Vec<u8>> {
        let archive = value.kind().archive().ok_or(SymbolError::Invalid(
            "the legend frame is stored by the chart drawable",
        ))?;
        if let Property::BackgroundFill(fill) = value {
            if matches!(fill, Fill::Image(_)) {
                return Err(SymbolError::Refused(Refusal::ImageFill));
            }
            let encoded = (*fill != default_background()).then(|| drawing::encode_fill(fill));
            let write = match encoded.as_deref() {
                Some(value) => Write::MergeMessage {
                    field: scalar::BACKGROUND_FILL,
                    value,
                    schema: codec::merge::schema::FILL,
                },
                None => Write::Message {
                    field: scalar::BACKGROUND_FILL,
                    value: None,
                },
            };
            return apply_writes(payload, archive, &[write], codec::EmptyExtension::Remove);
        }
        if let Property::Shadow(shadow) = *value {
            // The host's `patch_chart_shadow_scope`: the native default
            // removes the switch; otherwise it is written when grouped or
            // already stored.
            let present = flag(&decode_chart(payload, archive)?, scalar::COMBINE_LAYERS)?.is_some();
            let combined = shadow.groups_series();
            let stored = if shadow == Shadow::NATIVE_DEFAULT {
                None
            } else {
                (present || combined).then_some(combined)
            };
            return apply_writes(
                payload,
                archive,
                &[Write::Bool {
                    field: scalar::COMBINE_LAYERS,
                    value: stored,
                }],
                codec::EmptyExtension::Keep,
            );
        }
        if let Property::Lighting3d(lighting) = *value {
            let field =
                lighting_field(chart).ok_or(SymbolError::Refused(Refusal::UnsupportedKind))?;
            return apply_writes(
                payload,
                archive,
                &[Write::Message {
                    field,
                    value: lighting.package(),
                }],
                codec::EmptyExtension::Keep,
            );
        }
        if let Property::BorderStroke(stroke) = *value {
            let encoded =
                (stroke != Some(default_border())).then(|| drawing::encode_stroke(stroke));
            return apply_writes(
                payload,
                archive,
                &[Write::Message {
                    field: scalar::BORDER_STROKE,
                    value: encoded.as_deref(),
                }],
                codec::EmptyExtension::Keep,
            );
        }
        if let Some(writes) = scalar_writes(value.clone()) {
            return apply_writes(payload, archive, &writes, codec::EmptyExtension::Keep);
        }
        let write = rule(value.clone(), chart)?;
        apply_writes(payload, archive, write.writes(), write.empty())
    }

    /// The legend frame stored by a chart drawable payload.
    ///
    /// # Errors
    ///
    /// An invalid payload or stored frame.
    pub fn legend_frame(chart_drawable: &[u8]) -> SymbolResult<legend::Frame> {
        let stored =
            codec::legend_frame(chart_drawable, CodecLimits::DEFAULT).map_err(codec_error)?;
        legend::Frame::from_native(stored).map_err(out_of_range)
    }

    /// The chart drawable payload storing legend frame `frame`.
    ///
    /// # Errors
    ///
    /// An invalid payload.
    pub fn rewrite_legend_frame(
        chart_drawable: &[u8],
        frame: legend::Frame,
    ) -> SymbolResult<Vec<u8>> {
        Ok(
            codec::rewrite_legend_frame(chart_drawable, frame.native(), CodecLimits::DEFAULT)
                .map_err(codec_error)?
                .into_bytes()
                .into_owned(),
        )
    }
}

/// The axis-style properties of a chart's primary value or category axis.
pub mod axis {
    use litchi_iwa_common::chart::axis::style::Visibility;
    use litchi_iwa_common::chart::axis::{Axis, TickMarkLocation};
    use litchi_iwa_common::chart::kind::Kind as ChartKind;
    use litchi_iwa_common::wire;
    use litchi_iwa_protos::chart_style_codec::{
        self as codec, Archive, Limits as CodecLimits, Properties, Write,
    };
    use litchi_iwa_protos::shape_style_codec::native as stored;

    pub use litchi_iwa_common::chart::axis::label_angle::LabelAngle;
    pub use litchi_iwa_common::chart::axis::label_position_3d::LabelPosition3d;
    pub use litchi_iwa_common::chart::category_labels::{Frequency, Layout};
    pub use litchi_iwa_common::chart::number_format::{LabelAffixes, NumberFormat};
    pub use litchi_iwa_common::shape::stroke::Stroke;

    use super::{apply_writes, codec_error, out_of_range};
    use crate::drawing;
    use crate::graph::Error;
    use crate::number_format::{self as formats, Fields};
    use crate::series::{Refusal, SymbolError, SymbolResult};

    /// The axis-label formatter of `TSCH.ChartAxisNonStyleArchive` (the
    /// host's `AXIS_NUMBER_FORMAT_FIELDS`).
    const LABEL_FORMAT: Fields = Fields {
        legacy: 2,
        current: 42,
        format_type: 3,
    };

    /// The stroke override of one gridline family.
    #[derive(Debug, Clone, Copy, PartialEq, Default)]
    #[non_exhaustive]
    pub enum GridlineStroke {
        /// No override: the gridlines use the axis style's parent chain.
        #[default]
        Inherited,
        /// The inspector's "None" stroke.
        None,
        /// An explicit stroke.
        Stroke(Stroke),
    }

    /// One axis-style property with its value.
    #[derive(Debug, Clone, PartialEq)]
    #[non_exhaustive]
    pub enum Property {
        /// Whether the axis line shows.
        Line(Visibility),
        /// Whether major gridlines show.
        MajorGridlines(Visibility),
        /// Whether minor gridlines show.
        MinorGridlines(Visibility),
        /// Whether minor tick marks show.
        MinorTickMarks(Visibility),
        /// Where major tick marks are drawn.
        TickMarkLocation(TickMarkLocation),
        /// Whether the value axis shows its minimum label (value axis only).
        MinimumLabel(Visibility),
        /// Whether the axis labels show.
        Labels(Visibility),
        /// Whether the category axis shows the series names (category axis
        /// only).
        SeriesNames(Visibility),
        /// The stroke of the major gridlines.
        MajorGridlineStroke(GridlineStroke),
        /// The stroke of the minor gridlines.
        MinorGridlineStroke(GridlineStroke),
        /// The rotation of the axis labels.
        LabelAngle(LabelAngle),
        /// The prefix and suffix of the axis labels.
        LabelAffixes(LabelAffixes),
        /// The number format of the axis labels.
        NumberFormat(NumberFormat),
        /// Where a 3D chart draws its value-axis labels (value axis of the 3D
        /// kinds only).
        LabelPosition3d(LabelPosition3d),
        /// How often the category axis labels show and whether the last
        /// one always shows (category axis only). Hidden labels
        /// ([`Frequency::None`]) switch [`Property::Labels`] off and keep the
        /// stored interval.
        CategoryLabels(Layout),
    }

    /// The property a [`Property`] value sets, for reads.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    #[non_exhaustive]
    pub enum PropertyKind {
        /// [`Property::Line`].
        Line,
        /// [`Property::MajorGridlines`].
        MajorGridlines,
        /// [`Property::MinorGridlines`].
        MinorGridlines,
        /// [`Property::MinorTickMarks`].
        MinorTickMarks,
        /// [`Property::TickMarkLocation`].
        TickMarkLocation,
        /// [`Property::MinimumLabel`].
        MinimumLabel,
        /// [`Property::Labels`].
        Labels,
        /// [`Property::SeriesNames`].
        SeriesNames,
        /// [`Property::MajorGridlineStroke`].
        MajorGridlineStroke,
        /// [`Property::MinorGridlineStroke`].
        MinorGridlineStroke,
        /// [`Property::LabelAngle`].
        LabelAngle,
        /// [`Property::LabelAffixes`].
        LabelAffixes,
        /// [`Property::NumberFormat`].
        NumberFormat,
        /// [`Property::LabelPosition3d`].
        LabelPosition3d,
        /// [`Property::CategoryLabels`].
        CategoryLabels,
    }

    impl Property {
        /// The property this value sets.
        #[must_use]
        pub const fn kind(&self) -> PropertyKind {
            match self {
                Self::Line(_) => PropertyKind::Line,
                Self::MajorGridlines(_) => PropertyKind::MajorGridlines,
                Self::MinorGridlines(_) => PropertyKind::MinorGridlines,
                Self::MinorTickMarks(_) => PropertyKind::MinorTickMarks,
                Self::TickMarkLocation(_) => PropertyKind::TickMarkLocation,
                Self::MinimumLabel(_) => PropertyKind::MinimumLabel,
                Self::Labels(_) => PropertyKind::Labels,
                Self::SeriesNames(_) => PropertyKind::SeriesNames,
                Self::MajorGridlineStroke(_) => PropertyKind::MajorGridlineStroke,
                Self::MinorGridlineStroke(_) => PropertyKind::MinorGridlineStroke,
                Self::LabelAngle(_) => PropertyKind::LabelAngle,
                Self::LabelAffixes(_) => PropertyKind::LabelAffixes,
                Self::NumberFormat(_) => PropertyKind::NumberFormat,
                Self::LabelPosition3d(_) => PropertyKind::LabelPosition3d,
                Self::CategoryLabels(_) => PropertyKind::CategoryLabels,
            }
        }
    }

    impl PropertyKind {
        /// The extension field of the property on `axis` (the host's
        /// `axis_style`, `axis_gridline_stroke`, `axis_label_angle`,
        /// `axis_number_format` and `axis_label_position_3d` constants), or
        /// `None` when the axis has no such field. The label affixes and
        /// number format name the current formatter field.
        #[must_use]
        pub const fn field(self, axis: Axis) -> Option<u32> {
            Some(match (self, axis) {
                (Self::Line, Axis::Category) => 24,
                (Self::Line, Axis::Value) => 25,
                (Self::MajorGridlines, Axis::Category) => 27,
                (Self::MajorGridlines, Axis::Value) => 28,
                (Self::MinorGridlines, Axis::Category) => 32,
                (Self::MinorGridlines, Axis::Value) => 33,
                (Self::MinorTickMarks, Axis::Category) => 34,
                (Self::MinorTickMarks, Axis::Value) => 35,
                (Self::TickMarkLocation, Axis::Category) => 36,
                (Self::TickMarkLocation, Axis::Value) => 37,
                (Self::MinimumLabel, Axis::Value) => 31,
                (Self::Labels, Axis::Category) => 9,
                (Self::Labels, Axis::Value) => 11,
                (Self::SeriesNames, Axis::Category) => 12,
                (Self::MajorGridlineStroke, Axis::Category) => 16,
                (Self::MajorGridlineStroke, Axis::Value) => 17,
                (Self::MinorGridlineStroke, Axis::Category) => 22,
                (Self::MinorGridlineStroke, Axis::Value) => 23,
                (Self::LabelAngle, Axis::Category) => 9,
                (Self::LabelAngle, Axis::Value) => 11,
                (Self::LabelAffixes | Self::NumberFormat, _) => LABEL_FORMAT.current,
                (Self::LabelPosition3d, Axis::Value) => 1,
                (Self::CategoryLabels, Axis::Category) => CATEGORY_LABEL_INTERVAL,
                _ => return None,
            })
        }

        /// Whether the property exists on `axis` of a chart of kind `chart`.
        #[must_use]
        pub const fn applies_to(self, axis: Axis, chart: ChartKind) -> bool {
            self.field(axis).is_some()
                && (!matches!(self, Self::LabelPosition3d)
                    || chart.supports_3d_value_axis_label_position())
        }

        /// The value of a switch the style does not store.
        const fn default_visible(self) -> bool {
            matches!(
                self,
                Self::MinimumLabel | Self::MinorTickMarks | Self::Labels
            )
        }

        /// The axis archive storing the property.
        #[must_use]
        pub const fn archive(self) -> Archive {
            match self {
                Self::Labels
                | Self::SeriesNames
                | Self::LabelAffixes
                | Self::NumberFormat
                | Self::LabelPosition3d => Archive::AxisNonStyle,
                _ => Archive::AxisStyle,
            }
        }
    }

    /// `tschchartaxiscategorylabelsinterval` (category axis style).
    const CATEGORY_LABEL_INTERVAL: u32 = 5;
    /// `tschchartaxisshowlastlabel` (category axis style).
    const SHOW_LAST_CATEGORY: u32 = 26;

    fn invalid_labels<E>(_error: E) -> SymbolError {
        SymbolError::Invalid("an invalid category-label field")
    }

    /// A singular varint field of the extension, refusing a repeated field,
    /// another wire type or trailing bytes (the host's
    /// `strict_optional_varint`).
    fn strict_varint(extension: &[u8], field: u32) -> SymbolResult<Option<u64>> {
        use litchi_iwa_common::WireLimits;
        use litchi_iwa_common::varint::decode_varint_from_bytes;
        let view = wire::WireView::parse(extension, WireLimits::DEFAULT).map_err(invalid_labels)?;
        let mut found = None;
        for candidate in view.fields() {
            if candidate.number() != field {
                continue;
            }
            if found.is_some() || candidate.wire_type() != 0 {
                return Err(SymbolError::Invalid("an invalid category-label field"));
            }
            let payload = candidate.payload();
            let (value, consumed) = decode_varint_from_bytes(payload).map_err(invalid_labels)?;
            if consumed != payload.len() {
                return Err(SymbolError::Invalid("an invalid category-label field"));
            }
            found = Some(value);
        }
        Ok(found)
    }

    /// A stored signed 32-bit interval (a negative one is sign-extended).
    fn native_i32(raw: u64) -> SymbolResult<i32> {
        const NEGATIVE_INT32_START: u64 = 0xffff_ffff_8000_0000;
        if let Ok(value) = i32::try_from(raw) {
            return Ok(value);
        }
        if raw >= NEGATIVE_INT32_START {
            #[allow(
                clippy::cast_possible_truncation,
                clippy::cast_possible_wrap,
                reason = "a sign-extended int32 varint, as protobuf encodes it"
            )]
            return Ok(raw as i64 as i32);
        }
        Err(SymbolError::Invalid(
            "a category-label interval outside int32",
        ))
    }

    /// The stored interval and show-last switch of a decoded category axis
    /// style payload (absent: auto-fit, shown).
    ///
    /// # Errors
    ///
    /// A repeated or malformed field, a show-last value other than 0 or 1.
    pub fn category_label_style(properties: &Properties<'_>) -> SymbolResult<(Frequency, bool)> {
        let Some(extension) = properties.extension() else {
            return Ok((Frequency::AutoFit, true));
        };
        let frequency = match strict_varint(extension, CATEGORY_LABEL_INTERVAL)? {
            None | Some(0) => Frequency::AutoFit,
            Some(1) => Frequency::All,
            Some(raw) => Frequency::from_native(native_i32(raw)?),
        };
        let show_last = match strict_varint(extension, SHOW_LAST_CATEGORY)? {
            None | Some(1) => true,
            Some(0) => false,
            Some(_) => {
                return Err(SymbolError::Invalid(
                    "a show-last-category switch other than 0 or 1",
                ));
            },
        };
        Ok((frequency, show_last))
    }

    /// The category label layout from the decoded category axis style
    /// payload and the axis labels switch (the host's
    /// `chart_category_label_layout`).
    ///
    /// # Errors
    ///
    /// As [`category_label_style`].
    pub fn category_labels(
        properties: &Properties<'_>,
        labels: Visibility,
    ) -> SymbolResult<Layout> {
        let (frequency, show_last) = category_label_style(properties)?;
        Ok(Layout::new(
            if labels.is_visible() {
                frequency
            } else {
                Frequency::None
            },
            show_last,
        ))
    }

    /// The category axis style payload storing `frequency` (`None`: keep
    /// the stored interval) and the show-last switch, every other byte
    /// kept (the host's `patch_category_label_style`: auto-fit removes the
    /// interval; showing the last label keeps a stored switch and writes
    /// none otherwise; an emptied extension is removed).
    ///
    /// # Errors
    ///
    /// [`Frequency::None`] as the interval, an invalid payload.
    pub fn rewrite_category_labels(
        payload: &[u8],
        frequency: Option<Frequency>,
        show_last: bool,
    ) -> SymbolResult<Vec<u8>> {
        let wire_error = |error| SymbolError::Graph(Error::Wire(error));
        let properties = codec::decode(payload, Archive::AxisStyle, CodecLimits::DEFAULT)
            .map_err(codec_error)?;
        let existing = properties.extension();
        let extension = existing.unwrap_or_default();
        let interval_present = strict_varint(extension, CATEGORY_LABEL_INTERVAL)?.is_some();
        let show_last_present = strict_varint(extension, SHOW_LAST_CATEGORY)?.is_some();
        let mut patched = match frequency {
            Some(frequency) => {
                let replacement = match frequency {
                    Frequency::AutoFit => None,
                    Frequency::All => Some(1),
                    Frequency::Every(interval) => Some(u64::from(interval.value())),
                    #[allow(
                        clippy::cast_sign_loss,
                        reason = "a negative int32 is stored sign-extended"
                    )]
                    Frequency::Unsupported(value) => Some(i64::from(value) as u64),
                    _ => {
                        return Err(SymbolError::Invalid(
                            "hidden category labels have no stored interval",
                        ));
                    },
                };
                wire::patch_varint_field(
                    extension,
                    CATEGORY_LABEL_INTERVAL,
                    interval_present,
                    replacement,
                )
                .map_err(wire_error)?
            },
            None => extension.to_vec(),
        };
        let present = strict_varint(&patched, SHOW_LAST_CATEGORY)?.is_some();
        let replacement = if show_last {
            show_last_present.then_some(1)
        } else {
            Some(0)
        };
        patched = wire::patch_varint_field(&patched, SHOW_LAST_CATEGORY, present, replacement)
            .map_err(wire_error)?;
        if existing.is_none() && patched.is_empty() {
            return Ok(payload.to_vec());
        }
        wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            existing.is_some(),
            (!patched.is_empty()).then_some(patched.as_slice()),
        )
        .map_err(wire_error)
    }

    const fn visibility(visible: bool) -> Visibility {
        if visible {
            Visibility::Visible
        } else {
            Visibility::Hidden
        }
    }

    fn gridline_stroke(properties: &Properties<'_>, field: u32) -> SymbolResult<GridlineStroke> {
        let Some(bytes) = properties.message(field).map_err(codec_error)? else {
            return Ok(GridlineStroke::Inherited);
        };
        let decoded = stored::stroke(bytes, CodecLimits::DEFAULT).map_err(out_of_range)?;
        Ok(match drawing::stroke(&decoded).map_err(out_of_range)? {
            Some(stroke) => GridlineStroke::Stroke(stroke),
            None => GridlineStroke::None,
        })
    }

    fn format_error(error: Error) -> SymbolError {
        SymbolError::Graph(error)
    }

    /// The stored (or native default) value of `kind` on `axis` in the
    /// decoded payload of its [`PropertyKind::archive`].
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedAxis`], an invalid stored value.
    pub fn value(
        properties: &Properties<'_>,
        axis: Axis,
        kind: PropertyKind,
    ) -> SymbolResult<Property> {
        let field = kind
            .field(axis)
            .ok_or(SymbolError::Refused(Refusal::UnsupportedAxis))?;
        Ok(match kind {
            PropertyKind::TickMarkLocation => Property::TickMarkLocation(
                properties
                    .int32(field)
                    .map_err(codec_error)?
                    .map(TickMarkLocation::from_native)
                    .unwrap_or_default(),
            ),
            PropertyKind::MajorGridlineStroke => {
                Property::MajorGridlineStroke(gridline_stroke(properties, field)?)
            },
            PropertyKind::MinorGridlineStroke => {
                Property::MinorGridlineStroke(gridline_stroke(properties, field)?)
            },
            PropertyKind::LabelAngle => Property::LabelAngle(
                properties
                    .float(field)
                    .map_err(codec_error)?
                    .map(LabelAngle::new)
                    .transpose()
                    .map_err(out_of_range)?
                    .unwrap_or(LabelAngle::HORIZONTAL),
            ),
            PropertyKind::LabelAffixes => Property::LabelAffixes(
                formats::affixes(properties.extension(), LABEL_FORMAT).map_err(format_error)?,
            ),
            PropertyKind::NumberFormat => Property::NumberFormat(
                formats::number_format(
                    properties.extension(),
                    LABEL_FORMAT,
                    NumberFormat::AXIS_NATIVE_DEFAULT,
                )
                .map_err(format_error)?,
            ),
            PropertyKind::LabelPosition3d => Property::LabelPosition3d(
                properties
                    .int32(field)
                    .map_err(codec_error)?
                    .map(LabelPosition3d::from_native)
                    .unwrap_or_default(),
            ),
            PropertyKind::CategoryLabels => {
                return Err(SymbolError::Invalid(
                    "the category label layout also reads the labels switch",
                ));
            },
            _ => {
                let visible = visibility(
                    properties
                        .bool(field)
                        .map_err(codec_error)?
                        .unwrap_or(kind.default_visible()),
                );
                match kind {
                    PropertyKind::Line => Property::Line(visible),
                    PropertyKind::MajorGridlines => Property::MajorGridlines(visible),
                    PropertyKind::MinorGridlines => Property::MinorGridlines(visible),
                    PropertyKind::MinorTickMarks => Property::MinorTickMarks(visible),
                    PropertyKind::Labels => Property::Labels(visible),
                    PropertyKind::SeriesNames => Property::SeriesNames(visible),
                    _ => Property::MinimumLabel(visible),
                }
            },
        })
    }

    /// Rewrite the generated extension of `payload` with `patch` (`None`
    /// from the patch keeps the payload).
    fn patch_extension(
        payload: &[u8],
        archive: Archive,
        patch: impl FnOnce(&[u8]) -> crate::graph::Result<Option<Vec<u8>>>,
    ) -> SymbolResult<Vec<u8>> {
        let properties =
            codec::decode(payload, archive, CodecLimits::DEFAULT).map_err(codec_error)?;
        let existing = properties.extension();
        let Some(extension) = patch(existing.unwrap_or_default()).map_err(format_error)? else {
            return Ok(payload.to_vec());
        };
        wire::patch_length_delimited_field(
            payload,
            codec::EXTENSION_FIELD,
            existing.is_some(),
            Some(extension.as_slice()),
        )
        .map_err(|error| SymbolError::Graph(Error::Wire(error)))
    }

    /// The payload of the axis style object of `value`'s archive storing
    /// `value` on `axis`, every other byte kept.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedAxis`], an invalid payload.
    pub fn rewrite(payload: &[u8], axis: Axis, value: &Property) -> SymbolResult<Vec<u8>> {
        let kind = value.kind();
        let field = kind
            .field(axis)
            .ok_or(SymbolError::Refused(Refusal::UnsupportedAxis))?;
        let archive = kind.archive();
        let encoded;
        let (write, empty) = match value {
            Property::TickMarkLocation(location) => (
                Write::Int32 {
                    field,
                    value: Some(location.native_value()),
                },
                codec::EmptyExtension::Keep,
            ),
            Property::Line(visible)
            | Property::MajorGridlines(visible)
            | Property::MinorGridlines(visible)
            | Property::MinorTickMarks(visible)
            | Property::MinimumLabel(visible)
            | Property::Labels(visible)
            | Property::SeriesNames(visible) => (
                Write::Bool {
                    field,
                    value: Some(visible.is_visible()),
                },
                codec::EmptyExtension::Keep,
            ),
            Property::MajorGridlineStroke(stroke) | Property::MinorGridlineStroke(stroke) => {
                encoded = match stroke {
                    GridlineStroke::Inherited => None,
                    GridlineStroke::None => Some(drawing::encode_stroke(None)),
                    GridlineStroke::Stroke(stroke) => Some(drawing::encode_stroke(Some(*stroke))),
                };
                (
                    Write::Message {
                        field,
                        value: encoded.as_deref(),
                    },
                    codec::EmptyExtension::Keep,
                )
            },
            Property::LabelAngle(angle) => (
                Write::Float {
                    field,
                    value: (*angle != LabelAngle::HORIZONTAL).then_some(angle.degrees()),
                },
                codec::EmptyExtension::Remove,
            ),
            Property::LabelPosition3d(position) => {
                let properties =
                    codec::decode(payload, archive, CodecLimits::DEFAULT).map_err(codec_error)?;
                let present = properties.int32(field).map_err(codec_error)?.is_some();
                (
                    Write::Int32 {
                        field,
                        value: (present || *position != LabelPosition3d::default())
                            .then_some(position.native_value()),
                    },
                    codec::EmptyExtension::Keep,
                )
            },
            Property::LabelAffixes(affixes) => {
                return patch_extension(payload, archive, |extension| {
                    formats::patch_affixes(
                        extension,
                        LABEL_FORMAT,
                        affixes,
                        NumberFormat::AXIS_NATIVE_DEFAULT,
                    )
                });
            },
            Property::CategoryLabels(_) => {
                return Err(SymbolError::Invalid(
                    "the category label layout is written with rewrite_category_labels",
                ));
            },
            Property::NumberFormat(format) => {
                return patch_extension(payload, archive, |extension| {
                    if *format == NumberFormat::AXIS_NATIVE_DEFAULT {
                        formats::clear(extension, LABEL_FORMAT)
                    } else {
                        formats::patch_number_format(
                            extension,
                            LABEL_FORMAT,
                            *format,
                            NumberFormat::AXIS_NATIVE_DEFAULT,
                        )
                    }
                });
            },
        };
        apply_writes(payload, archive, &[write], empty)
    }
}

/// The chart's style objects of `archive` on `axis`.
fn axis_objects(references: &StyleReferences, axis: Axis, archive: Archive) -> &[u64] {
    match (axis, archive) {
        (Axis::Value, Archive::AxisNonStyle) => references.value_axis_non_styles(),
        (_, Archive::AxisNonStyle) => references.category_axis_non_styles(),
        (Axis::Value, _) => references.value_axis_styles(),
        _ => references.category_axis_styles(),
    }
}

/// How many chart references in the workspace name `style` in the role of
/// `archive` (chart-level, or any axis).
pub(crate) fn owners(
    workspace: &Workspace<'_>,
    style: u64,
    archive: Archive,
    budget: &mut SymbolBudget,
) -> SymbolResult<usize> {
    let mut count = 0_usize;
    for location in workspace.locations_iter() {
        let components = workspace.archive_at(location).map_err(SymbolError::from)?;
        for object in &components.objects {
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let references = bounded_references(&message.data, budget)?;
                count = count.saturating_add(match archive {
                    Archive::AxisStyle | Archive::AxisNonStyle => {
                        axis_objects(&references, Axis::Value, archive)
                            .iter()
                            .chain(axis_objects(&references, Axis::Category, archive))
                            .filter(|identifier| **identifier == style)
                            .count()
                    },
                    _ => usize::from(references.chart_level(archive) == Some(style)),
                });
            }
        }
    }
    Ok(count)
}

/// The style object of `archive` the chart (on `axis`) references.
pub(crate) fn style_object(
    workspace: &Workspace<'_>,
    chart: u64,
    archive: Archive,
    axis: Option<Axis>,
    budget: &mut SymbolBudget,
) -> SymbolResult<u64> {
    let (_, payload) = bounded_chart_payload(workspace, chart, budget)?;
    let references = bounded_references(payload, budget)?;
    match axis {
        Some(axis) => axis_objects(&references, axis, archive)
            .first()
            .copied()
            .ok_or(SymbolError::Refused(Refusal::UnsupportedAxis)),
        None => references
            .chart_level(archive)
            .ok_or(SymbolError::Refused(Refusal::MissingStyle)),
    }
}

/// The in-place scope of rewriting `object` (messages of `message_type`).
pub(crate) fn in_place_scope(
    workspace: &Workspace<'_>,
    chart: u64,
    object: u64,
    message_type: u32,
    chart_fields: &[u32],
    budget: &mut SymbolBudget,
) -> SymbolResult<()> {
    let location = workspace.locate(object).map_err(SymbolError::from)?;
    let member = workspace.member(location);
    budget.preflight_allocations(4)?;
    budget.preflight_retained(member.len().saturating_add(32))?;
    budget.allocations(4)?;
    budget.retained(member.len().saturating_add(32))?;
    budget.set_scope(MutationScope::in_place(
        chart,
        vec![object],
        vec![member.to_owned()],
        vec![message_type],
        vec![codec::EXTENSION_FIELD],
        chart_fields.to_vec(),
    ));
    Ok(())
}

/// Rewrite the unique `message_type` payload of `object` with `rewrite`,
/// returning whether it changed.
pub(crate) fn rewrite_in_place(
    workspace: &mut Workspace<'_>,
    object: u64,
    message_type: u32,
    budget: &mut SymbolBudget,
    rewrite: impl FnOnce(&[u8]) -> SymbolResult<Vec<u8>>,
) -> SymbolResult<bool> {
    let (message, payload) = bounded_payload(
        workspace,
        object,
        message_type,
        "a chart style object must have exactly one payload",
        budget,
    )?;
    let data = charged_patch(payload, budget, rewrite)?;
    if data == payload {
        return Ok(false);
    }
    workspace
        .replace_payload(object, message, data)
        .map_err(SymbolError::from)?;
    Ok(true)
}

pub(crate) fn read_payload<T>(
    workspace: &Workspace<'_>,
    object: u64,
    message_type: u32,
    budget: &mut SymbolBudget,
    read: impl FnOnce(&[u8]) -> SymbolResult<T>,
) -> SymbolResult<T> {
    let (_, payload) = bounded_payload(
        workspace,
        object,
        message_type,
        "a chart style object must have exactly one payload",
        budget,
    )?;
    budget.source(payload.len())?;
    budget.work(payload.len())?;
    read(payload)
}

fn chart_value(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    series_count: usize,
    kind: chart::PropertyKind,
    budget: &mut SymbolBudget,
) -> SymbolResult<chart::Property> {
    if !kind.applies_to(chart_kind) {
        return Err(SymbolError::Refused(Refusal::UnsupportedKind));
    }
    if kind == chart::PropertyKind::RadarSeriesStyle {
        return crate::series::style::radar::read(workspace, chart, series_count, budget)
            .map(chart::Property::RadarSeriesStyle);
    }
    let Some(archive) = kind.archive() else {
        return read_payload(workspace, chart, CHART_MESSAGE_TYPE, budget, |payload| {
            chart::legend_frame(payload).map(chart::Property::LegendFrame)
        });
    };
    let style = style_object(workspace, chart, archive, None, budget)?;
    if kind == chart::PropertyKind::Shadow {
        let series = crate::series::style::chart_shadow(workspace, chart, budget)?;
        return read_payload(
            workspace,
            style,
            archive.message_type(),
            budget,
            |payload| {
                chart::shadow(&decode(payload, archive)?, series).map(chart::Property::Shadow)
            },
        );
    }
    read_payload(
        workspace,
        style,
        archive.message_type(),
        budget,
        |payload| chart::value(&decode(payload, archive)?, kind, chart_kind),
    )
}

/// Read chart-level property `kind` of chart drawable `chart` (of kind
/// `chart_kind`, with `series_count` visible series) under one residual
/// profile.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], [`Refusal::MissingStyle`], an invalid
/// chart or style, or an exhausted limit.
pub fn read_chart(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    series_count: usize,
    kind: chart::PropertyKind,
    limits: Limits,
) -> SymbolResult<(chart::Property, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let value = chart_value(
        workspace,
        chart,
        chart_kind,
        series_count,
        kind,
        &mut budget,
    )?;
    Ok((value, budget.report(false)))
}

/// Store chart-level `value` for chart drawable `chart` (of kind
/// `chart_kind`, with `series_count` visible series) under one residual
/// profile: the style object is rewritten in place (the legend frame: the
/// chart drawable; the radar series style: the series styles of the visible
/// series). Returns whether anything changed and the report with the
/// in-place mutation scope.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], [`Refusal::MissingStyle`],
/// [`Refusal::SharedStyle`], an invalid chart or style, an exhausted limit,
/// or a write that does not read back.
pub fn write_chart(
    workspace: &mut Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    series_count: usize,
    value: &chart::Property,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let kind = value.kind();
    if chart_value(
        workspace,
        chart,
        chart_kind,
        series_count,
        kind,
        &mut budget,
    )? == *value
    {
        return Ok((false, budget.report(false)));
    }
    let changed = match (kind.archive(), value) {
        (_, chart::Property::RadarSeriesStyle(style)) => {
            // The host's `set_chart_radar_series_style` checks no chart
            // style; the scope is the chart's series graph.
            let scope = crate::series::radar_scope(workspace, chart, series_count, &mut budget)?;
            budget.set_scope(scope);
            crate::series::style::radar::write(workspace, chart, series_count, *style, &mut budget)?
        },
        (None, _) => {
            in_place_scope(
                workspace,
                chart,
                chart,
                CHART_MESSAGE_TYPE,
                &[LEGEND_FRAME_FIELD],
                &mut budget,
            )?;
            let chart::Property::LegendFrame(frame) = *value else {
                return Err(SymbolError::Invalid("not a legend frame"));
            };
            rewrite_in_place(
                workspace,
                chart,
                CHART_MESSAGE_TYPE,
                &mut budget,
                |payload| chart::rewrite_legend_frame(payload, frame),
            )?
        },
        (Some(archive), _) => {
            let style = style_object(workspace, chart, archive, None, &mut budget)?;
            if owners(workspace, style, archive, &mut budget)? != 1 {
                return Err(SymbolError::Refused(Refusal::SharedStyle));
            }
            in_place_scope(
                workspace,
                chart,
                style,
                archive.message_type(),
                &[],
                &mut budget,
            )?;
            let series = match *value {
                // The host's `set_chart_shadow`: every series style is
                // exclusive and rewritten before the chart style.
                chart::Property::Shadow(shadow) => crate::series::style::write_chart_shadow(
                    workspace,
                    chart,
                    shadow
                        .drop_shadow()
                        .map_or(StoredShadow::Disabled, StoredShadow::Drop),
                    &mut budget,
                )?,
                _ => false,
            };
            let style_changed = rewrite_in_place(
                workspace,
                style,
                archive.message_type(),
                &mut budget,
                |payload| chart::rewrite(payload, value, chart_kind),
            )?;
            series || style_changed
        },
    };
    if changed
        && chart_value(
            workspace,
            chart,
            chart_kind,
            series_count,
            kind,
            &mut budget,
        )? != *value
    {
        return Err(SymbolError::Invalid(
            "a chart style write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}

fn axis_value(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    axis: Axis,
    kind: axis::PropertyKind,
    budget: &mut SymbolBudget,
) -> SymbolResult<axis::Property> {
    if kind.field(axis).is_none() {
        return Err(SymbolError::Refused(Refusal::UnsupportedAxis));
    }
    if !kind.applies_to(axis, chart_kind) {
        return Err(SymbolError::Refused(Refusal::UnsupportedKind));
    }
    let archive = kind.archive();
    let style = style_object(workspace, chart, archive, Some(axis), budget)?;
    if kind == axis::PropertyKind::CategoryLabels {
        let axis::Property::Labels(labels) = axis_value(
            workspace,
            chart,
            chart_kind,
            axis,
            axis::PropertyKind::Labels,
            budget,
        )?
        else {
            return Err(SymbolError::Invalid("not an axis labels switch"));
        };
        return read_payload(
            workspace,
            style,
            archive.message_type(),
            budget,
            |payload| {
                axis::category_labels(&decode(payload, archive)?, labels)
                    .map(axis::Property::CategoryLabels)
            },
        );
    }
    read_payload(
        workspace,
        style,
        archive.message_type(),
        budget,
        |payload| axis::value(&decode(payload, archive)?, axis, kind),
    )
}

/// The host's `set_chart_category_label_layout`: the category axis style
/// is rewritten (after its exclusivity check) only when the show-last
/// switch or a shown interval changes, then the labels switch follows the
/// layout's visibility through the ordinary labels write.
fn write_category_labels(
    workspace: &mut Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    axis: Axis,
    layout: axis::Layout,
    budget: &mut SymbolBudget,
) -> SymbolResult<bool> {
    let style = style_object(workspace, chart, Archive::AxisStyle, Some(axis), budget)?;
    in_place_scope(
        workspace,
        chart,
        style,
        Archive::AxisStyle.message_type(),
        &[],
        budget,
    )?;
    let (stored, show_last) = read_payload(
        workspace,
        style,
        Archive::AxisStyle.message_type(),
        budget,
        |payload| axis::category_label_style(&decode(payload, Archive::AxisStyle)?),
    )?;
    let requested = match layout.frequency() {
        axis::Frequency::None => None,
        frequency => Some(frequency),
    };
    let mut changed = false;
    if show_last != layout.show_last_category()
        || requested.is_some_and(|frequency| frequency != stored)
    {
        if owners(workspace, style, Archive::AxisStyle, budget)? != 1 {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
        changed |= rewrite_in_place(
            workspace,
            style,
            Archive::AxisStyle.message_type(),
            budget,
            |payload| {
                axis::rewrite_category_labels(payload, requested, layout.show_last_category())
            },
        )?;
    }
    let labels = axis::Property::Labels(if layout.frequency().is_visible() {
        Visibility::Visible
    } else {
        Visibility::Hidden
    });
    if axis_value(
        workspace,
        chart,
        chart_kind,
        axis,
        axis::PropertyKind::Labels,
        budget,
    )? != labels
    {
        let non_style = style_object(workspace, chart, Archive::AxisNonStyle, Some(axis), budget)?;
        if owners(workspace, non_style, Archive::AxisNonStyle, budget)? != 1 {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
        let location = workspace.locate(non_style).map_err(SymbolError::from)?;
        let member = workspace.member(location).to_owned();
        if let Some(scope) = budget.scope_mut() {
            scope.allow(non_style, &member, Archive::AxisNonStyle.message_type());
        }
        changed |= rewrite_in_place(
            workspace,
            non_style,
            Archive::AxisNonStyle.message_type(),
            budget,
            |payload| axis::rewrite(payload, axis, &labels),
        )?;
    }
    Ok(changed)
}

/// Read axis property `kind` of the primary `axis` of chart drawable
/// `chart` (of kind `chart_kind`) under one residual profile.
///
/// # Errors
///
/// [`Refusal::UnsupportedAxis`], [`Refusal::UnsupportedKind`], an invalid
/// chart or style, or an exhausted limit.
pub fn read_axis(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    axis: Axis,
    kind: axis::PropertyKind,
    limits: Limits,
) -> SymbolResult<(axis::Property, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let value = axis_value(workspace, chart, chart_kind, axis, kind, &mut budget)?;
    Ok((value, budget.report(false)))
}

/// Store axis `value` on the chart's primary `axis` under one residual
/// profile, rewriting the axis style in place when only this chart's axes
/// use it.
///
/// # Errors
///
/// [`Refusal::UnsupportedAxis`], [`Refusal::UnsupportedKind`],
/// [`Refusal::SharedStyle`], an invalid chart or style, an exhausted limit,
/// or a write that does not read back.
pub fn write_axis(
    workspace: &mut Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    axis: Axis,
    value: &axis::Property,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let kind = value.kind();
    if axis_value(workspace, chart, chart_kind, axis, kind, &mut budget)? == *value {
        return Ok((false, budget.report(false)));
    }
    let archive = kind.archive();
    let changed = if let axis::Property::CategoryLabels(layout) = *value {
        write_category_labels(workspace, chart, chart_kind, axis, layout, &mut budget)?
    } else {
        let style = style_object(workspace, chart, archive, Some(axis), &mut budget)?;
        if owners(workspace, style, archive, &mut budget)? != 1 {
            return Err(SymbolError::Refused(Refusal::SharedStyle));
        }
        in_place_scope(
            workspace,
            chart,
            style,
            archive.message_type(),
            &[],
            &mut budget,
        )?;
        rewrite_in_place(
            workspace,
            style,
            archive.message_type(),
            &mut budget,
            |payload| axis::rewrite(payload, axis, value),
        )?
    };
    if changed && axis_value(workspace, chart, chart_kind, axis, kind, &mut budget)? != *value {
        return Err(SymbolError::Invalid(
            "an axis style write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}
