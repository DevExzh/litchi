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
use litchi_iwa_common::chart::kind::Kind as ChartKind;
use litchi_iwa_protos::chart_style_codec::{
    self as codec, Archive, Limits as CodecLimits, Properties, StyleReferences,
};

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

/// The chart-level properties.
pub mod chart {
    use litchi_iwa_common::chart::kind::Kind as ChartKind;
    use litchi_iwa_common::chart::{corner, donut, gaps, legend, pie, radar};
    use litchi_iwa_common::chart3d::{self, Scale};
    use litchi_iwa_protos::chart_style_codec::{
        self as codec, Archive, Limits as CodecLimits, Properties, Write, field, property,
    };

    use super::{apply_writes, codec_error, out_of_range};
    use crate::series::{Refusal, SymbolError, SymbolResult};

    /// One chart-level style property with its value.
    #[derive(Debug, Clone, Copy, PartialEq)]
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
            }
        }
    }

    impl PropertyKind {
        /// The native default of this property (what a reset writes).
        #[must_use]
        pub const fn native_default(self) -> Property {
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
            }
        }

        /// The style archive storing the property; `None` for the legend
        /// frame, which the chart drawable stores.
        #[must_use]
        pub const fn archive(self) -> Option<Archive> {
            Some(match self {
                Self::LegendFrame => return None,
                Self::RoundedCorners
                | Self::BorderVisible
                | Self::RadarGridShape
                | Self::GapSpacing => Archive::ChartStyle,
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
                Self::RoundedCorners
                | Self::LegendFrame
                | Self::LegendVisible
                | Self::BorderVisible
                | Self::IncludesHiddenData
                | Self::GapSpacing => true,
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
            PropertyKind::LegendFrame => {
                return Err(SymbolError::Invalid(
                    "the legend frame is stored by the chart drawable",
                ));
            },
        })
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
    pub fn rewrite(payload: &[u8], value: Property, chart: ChartKind) -> SymbolResult<Vec<u8>> {
        let archive = value.kind().archive().ok_or(SymbolError::Invalid(
            "the legend frame is stored by the chart drawable",
        ))?;
        if let Some(writes) = scalar_writes(value) {
            return apply_writes(payload, archive, &writes, codec::EmptyExtension::Keep);
        }
        let write = rule(value, chart)?;
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
    use litchi_iwa_protos::chart_style_codec::{self as codec, Archive, Properties, Write};

    use super::{apply_writes, codec_error};
    use crate::series::{Refusal, SymbolError, SymbolResult};

    /// One axis-style property with its value.
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
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
            }
        }
    }

    impl PropertyKind {
        /// The extension field of the property on `axis` (the host's
        /// `axis_style` constants), or `None` when the axis has no such
        /// field.
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
                _ => return None,
            })
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
                Self::Labels | Self::SeriesNames => Archive::AxisNonStyle,
                _ => Archive::AxisStyle,
            }
        }
    }

    const fn visibility(visible: bool) -> Visibility {
        if visible {
            Visibility::Visible
        } else {
            Visibility::Hidden
        }
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
        if kind == PropertyKind::TickMarkLocation {
            return Ok(Property::TickMarkLocation(
                properties
                    .int32(field)
                    .map_err(codec_error)?
                    .map(TickMarkLocation::from_native)
                    .unwrap_or_default(),
            ));
        }
        let visible = visibility(
            properties
                .bool(field)
                .map_err(codec_error)?
                .unwrap_or(kind.default_visible()),
        );
        Ok(match kind {
            PropertyKind::Line => Property::Line(visible),
            PropertyKind::MajorGridlines => Property::MajorGridlines(visible),
            PropertyKind::MinorGridlines => Property::MinorGridlines(visible),
            PropertyKind::MinorTickMarks => Property::MinorTickMarks(visible),
            PropertyKind::Labels => Property::Labels(visible),
            PropertyKind::SeriesNames => Property::SeriesNames(visible),
            _ => Property::MinimumLabel(visible),
        })
    }

    /// The payload of the axis style object of `value`'s archive storing
    /// `value` on `axis`, every other byte kept.
    ///
    /// # Errors
    ///
    /// [`Refusal::UnsupportedAxis`], an invalid payload.
    pub fn rewrite(payload: &[u8], axis: Axis, value: Property) -> SymbolResult<Vec<u8>> {
        let kind = value.kind();
        let field = kind
            .field(axis)
            .ok_or(SymbolError::Refused(Refusal::UnsupportedAxis))?;
        let write = match value {
            Property::TickMarkLocation(location) => Write::Int32 {
                field,
                value: Some(location.native_value()),
            },
            Property::Line(visible)
            | Property::MajorGridlines(visible)
            | Property::MinorGridlines(visible)
            | Property::MinorTickMarks(visible)
            | Property::MinimumLabel(visible)
            | Property::Labels(visible)
            | Property::SeriesNames(visible) => Write::Bool {
                field,
                value: Some(visible.is_visible()),
            },
        };
        apply_writes(
            payload,
            kind.archive(),
            &[write],
            codec::EmptyExtension::Keep,
        )
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
fn owners(
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
fn style_object(
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
fn in_place_scope(
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
fn rewrite_in_place(
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

fn read_payload<T>(
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
    kind: chart::PropertyKind,
    budget: &mut SymbolBudget,
) -> SymbolResult<chart::Property> {
    if !kind.applies_to(chart_kind) {
        return Err(SymbolError::Refused(Refusal::UnsupportedKind));
    }
    let Some(archive) = kind.archive() else {
        return read_payload(workspace, chart, CHART_MESSAGE_TYPE, budget, |payload| {
            chart::legend_frame(payload).map(chart::Property::LegendFrame)
        });
    };
    let style = style_object(workspace, chart, archive, None, budget)?;
    read_payload(
        workspace,
        style,
        archive.message_type(),
        budget,
        |payload| chart::value(&decode(payload, archive)?, kind, chart_kind),
    )
}

/// Read chart-level property `kind` of chart drawable `chart` (of kind
/// `chart_kind`) under one residual profile.
///
/// # Errors
///
/// [`Refusal::UnsupportedKind`], [`Refusal::MissingStyle`], an invalid
/// chart or style, or an exhausted limit.
pub fn read_chart(
    workspace: &Workspace<'_>,
    chart: u64,
    chart_kind: ChartKind,
    kind: chart::PropertyKind,
    limits: Limits,
) -> SymbolResult<(chart::Property, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let value = chart_value(workspace, chart, chart_kind, kind, &mut budget)?;
    Ok((value, budget.report(false)))
}

/// Store chart-level `value` for chart drawable `chart` (of kind
/// `chart_kind`) under one residual profile: the style object is rewritten
/// in place (the legend frame: the chart drawable). Returns whether
/// anything changed and the report with the in-place mutation scope.
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
    value: chart::Property,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let kind = value.kind();
    if chart_value(workspace, chart, chart_kind, kind, &mut budget)? == value {
        return Ok((false, budget.report(false)));
    }
    let changed = match kind.archive() {
        None => {
            in_place_scope(
                workspace,
                chart,
                chart,
                CHART_MESSAGE_TYPE,
                &[LEGEND_FRAME_FIELD],
                &mut budget,
            )?;
            let chart::Property::LegendFrame(frame) = value else {
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
        Some(archive) => {
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
            rewrite_in_place(
                workspace,
                style,
                archive.message_type(),
                &mut budget,
                |payload| chart::rewrite(payload, value, chart_kind),
            )?
        },
    };
    if changed && chart_value(workspace, chart, chart_kind, kind, &mut budget)? != value {
        return Err(SymbolError::Invalid(
            "a chart style write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}

fn axis_value(
    workspace: &Workspace<'_>,
    chart: u64,
    axis: Axis,
    kind: axis::PropertyKind,
    budget: &mut SymbolBudget,
) -> SymbolResult<axis::Property> {
    if kind.field(axis).is_none() {
        return Err(SymbolError::Refused(Refusal::UnsupportedAxis));
    }
    let archive = kind.archive();
    let style = style_object(workspace, chart, archive, Some(axis), budget)?;
    read_payload(
        workspace,
        style,
        archive.message_type(),
        budget,
        |payload| axis::value(&decode(payload, archive)?, axis, kind),
    )
}

/// Read axis property `kind` of the chart's primary `axis` under one
/// residual profile.
///
/// # Errors
///
/// [`Refusal::UnsupportedAxis`], an invalid chart or style, or an
/// exhausted limit.
pub fn read_axis(
    workspace: &Workspace<'_>,
    chart: u64,
    axis: Axis,
    kind: axis::PropertyKind,
    limits: Limits,
) -> SymbolResult<(axis::Property, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let value = axis_value(workspace, chart, axis, kind, &mut budget)?;
    Ok((value, budget.report(false)))
}

/// Store axis `value` on the chart's primary `axis` under one residual
/// profile, rewriting the axis style in place when only this chart's axes
/// use it.
///
/// # Errors
///
/// [`Refusal::UnsupportedAxis`], [`Refusal::SharedStyle`], an invalid chart
/// or style, an exhausted limit, or a write that does not read back.
pub fn write_axis(
    workspace: &mut Workspace<'_>,
    chart: u64,
    axis: Axis,
    value: axis::Property,
    limits: Limits,
) -> SymbolResult<(bool, Report)> {
    let mut budget = SymbolBudget::new(limits);
    let kind = value.kind();
    if axis_value(workspace, chart, axis, kind, &mut budget)? == value {
        return Ok((false, budget.report(false)));
    }
    let archive = kind.archive();
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
    let changed = rewrite_in_place(
        workspace,
        style,
        archive.message_type(),
        &mut budget,
        |payload| axis::rewrite(payload, axis, value),
    )?;
    if changed && axis_value(workspace, chart, axis, kind, &mut budget)? != value {
        return Err(SymbolError::Invalid(
            "an axis style write does not read back",
        ));
    }
    Ok((changed, budget.report(changed)))
}
