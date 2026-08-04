//! Lossless native chart-legend frame CRUD.
//!
//! iWork stores an optional legend rectangle directly in `TSCH.ChartArchive`.
//! Absence delegates placement and measurement to the chart layout engine;
//! native applications may materialize a rectangle whose zero extents retain
//! automatic measurement while fixing the legend origin.

use prost::Message;

use crate::archive::RawMessage;
use crate::charts::source::CHART_MESSAGE_TYPE;
use crate::protobuf::{tsch, tsp};
use crate::wire::{parse_wire_fields, patch_length_delimited_field};
use crate::{Error, IWorkPackage, Result};

const CHART_ARCHIVE_EXTENSION_FIELD: u32 = 10_000;
const CHART_LEGEND_FRAME_FIELD: u32 = 3;
const RECT_ORIGIN_FIELD: u32 = 1;
const RECT_SIZE_FIELD: u32 = 2;
const POINT_X_FIELD: u32 = 1;
const POINT_Y_FIELD: u32 = 2;
const SIZE_WIDTH_FIELD: u32 = 1;
const SIZE_HEIGHT_FIELD: u32 = 2;

/// A finite point-valued legend-frame coordinate in native chart units.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartLegendCoordinate(f32);

impl ChartLegendCoordinate {
    /// Coordinate zero.
    pub const ZERO: Self = Self(0.0);

    /// Construct a finite coordinate.
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::InvalidFormat(
                "chart legend coordinate must be finite".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the coordinate in native chart units.
    pub const fn points(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for ChartLegendCoordinate {
    type Error = Error;

    fn try_from(points: f32) -> Result<Self> {
        Self::new(points)
    }
}

/// A finite, non-negative legend-frame extent in native chart units.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ChartLegendExtent(f32);

impl ChartLegendExtent {
    /// A zero extent asks iWork to measure that axis automatically.
    pub const AUTOMATIC: Self = Self(0.0);

    /// Construct a finite, non-negative extent.
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::InvalidFormat(
                "chart legend extent must be finite and non-negative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    /// Return the extent in native chart units.
    pub const fn points(self) -> f32 {
        self.0
    }

    /// Return whether iWork should measure this axis automatically.
    pub const fn is_automatic(self) -> bool {
        self.0 == 0.0
    }
}

impl TryFrom<f32> for ChartLegendExtent {
    type Error = Error;

    fn try_from(points: f32) -> Result<Self> {
        Self::new(points)
    }
}

/// An explicit native legend rectangle.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ChartLegendRect {
    x: ChartLegendCoordinate,
    y: ChartLegendCoordinate,
    width: ChartLegendExtent,
    height: ChartLegendExtent,
}

impl ChartLegendRect {
    /// Construct a rectangle from strictly typed native chart units.
    pub const fn new(
        x: ChartLegendCoordinate,
        y: ChartLegendCoordinate,
        width: ChartLegendExtent,
        height: ChartLegendExtent,
    ) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Validate and construct a rectangle from raw native chart units.
    pub fn from_points(x: f32, y: f32, width: f32, height: f32) -> Result<Self> {
        Ok(Self::new(
            ChartLegendCoordinate::new(x)?,
            ChartLegendCoordinate::new(y)?,
            ChartLegendExtent::new(width)?,
            ChartLegendExtent::new(height)?,
        ))
    }

    /// Return the horizontal origin.
    #[must_use]
    pub const fn x(self) -> ChartLegendCoordinate {
        self.x
    }

    /// Return the vertical origin.
    #[must_use]
    pub const fn y(self) -> ChartLegendCoordinate {
        self.y
    }

    /// Return the width.
    #[must_use]
    pub const fn width(self) -> ChartLegendExtent {
        self.width
    }

    /// Return the height.
    #[must_use]
    pub const fn height(self) -> ChartLegendExtent {
        self.height
    }
}

/// Exact direct chart-legend frame state.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ChartLegendFrame {
    /// Let the native chart layout engine place and measure the legend.
    #[default]
    Automatic,
    /// Persist an explicit origin and optional fixed extents.
    Frame(ChartLegendRect),
}

/// Read one chart's exact native legend-frame state.
pub(crate) fn chart_legend_frame(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<ChartLegendFrame> {
    let (_, data) = chart_message(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    read_chart_legend_frame(&data)
}

/// Set or remove one chart's native legend rectangle.
pub(crate) fn set_chart_legend_frame(
    package: &mut IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
    frame: ChartLegendFrame,
) -> Result<()> {
    let (message_index, data) = chart_message(
        package,
        chart_archive_name,
        drawable_object_id,
        drawable_label,
    )?;
    let current = read_chart_legend_frame(&data)?;
    if current == frame {
        return Ok(());
    }
    let chart = chart_payload(&data)?;
    let frame_present =
        strict_optional_payload(chart, CHART_LEGEND_FRAME_FIELD, "legend frame")?.is_some();
    let encoded = match frame {
        ChartLegendFrame::Automatic => None,
        ChartLegendFrame::Frame(rect) => Some(rect_to_native(rect).encode_to_vec()),
    };
    let chart = patch_length_delimited_field(
        chart,
        CHART_LEGEND_FRAME_FIELD,
        frame_present,
        encoded.as_deref(),
    )?;
    let data =
        patch_length_delimited_field(&data, CHART_ARCHIVE_EXTENSION_FIELD, true, Some(&chart))?;
    if read_chart_legend_frame(&data)? != frame {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} legend-frame update failed validation"
        )));
    }
    package.update_archive(chart_archive_name, |archive| {
        let object = archive.object_mut(drawable_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "{drawable_label} chart {drawable_object_id} is missing"
            ))
        })?;
        Ok(object
            .replace_message(
                message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data,
                },
            )
            .map(|_| ())?)
    })
}

fn read_chart_legend_frame(data: &[u8]) -> Result<ChartLegendFrame> {
    let chart = chart_payload(data)?;
    let Some(frame) = strict_optional_payload(chart, CHART_LEGEND_FRAME_FIELD, "legend frame")?
    else {
        return Ok(ChartLegendFrame::Automatic);
    };
    strict_rect(frame).map(ChartLegendFrame::Frame)
}

fn rect_to_native(rect: ChartLegendRect) -> tsch::RectArchive {
    tsch::RectArchive {
        origin: tsp::Point {
            x: rect.x.points(),
            y: rect.y.points(),
        },
        size: tsp::Size {
            width: rect.width.points(),
            height: rect.height.points(),
        },
    }
}

fn strict_rect(data: &[u8]) -> Result<ChartLegendRect> {
    let decoded = tsch::RectArchive::decode(data)?;
    let origin = strict_required_payload(data, RECT_ORIGIN_FIELD, "legend frame origin")?;
    let size = strict_required_payload(data, RECT_SIZE_FIELD, "legend frame size")?;
    strict_required_fixed32(origin, POINT_X_FIELD, "legend frame x")?;
    strict_required_fixed32(origin, POINT_Y_FIELD, "legend frame y")?;
    strict_required_fixed32(size, SIZE_WIDTH_FIELD, "legend frame width")?;
    strict_required_fixed32(size, SIZE_HEIGHT_FIELD, "legend frame height")?;
    ChartLegendRect::from_points(
        decoded.origin.x,
        decoded.origin.y,
        decoded.size.width,
        decoded.size.height,
    )
}

fn chart_payload(data: &[u8]) -> Result<&[u8]> {
    strict_required_payload(
        data,
        CHART_ARCHIVE_EXTENSION_FIELD,
        "chart archive extension",
    )
}

fn strict_required_payload<'a>(data: &'a [u8], field_number: u32, label: &str) -> Result<&'a [u8]> {
    strict_optional_payload(data, field_number, label)?
        .ok_or_else(|| Error::InvalidFormat(format!("{label} field {field_number} is missing")))
}

fn strict_optional_payload<'a>(
    data: &'a [u8],
    field_number: u32,
    label: &str,
) -> Result<Option<&'a [u8]>> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() > 1 {
        return Err(Error::InvalidFormat(format!(
            "{label} field occurs {} times",
            matches.len()
        )));
    }
    let Some(field) = matches.first() else {
        return Ok(None);
    };
    if field.wire_type != 2 {
        return Err(Error::InvalidFormat(format!(
            "{label} field must be length-delimited"
        )));
    }
    Ok(Some(&data[field.payload_start..field.end]))
}

fn strict_required_fixed32(data: &[u8], field_number: u32, label: &str) -> Result<()> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number)
        .collect::<Vec<_>>();
    if matches.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "{label} field must occur exactly once"
        )));
    }
    if matches[0].wire_type != 5 {
        return Err(Error::InvalidFormat(format!(
            "{label} field must be fixed32"
        )));
    }
    Ok(())
}

fn chart_message(
    package: &IWorkPackage,
    chart_archive_name: &str,
    drawable_object_id: u64,
    drawable_label: &str,
) -> Result<(usize, Vec<u8>)> {
    let archive = package.archive(chart_archive_name)?;
    let object = archive.object(drawable_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} is missing"
        ))
    })?;
    let messages = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == CHART_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [(message_index, message)] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{drawable_label} chart {drawable_object_id} must have exactly one chart payload"
        )));
    };
    Ok((*message_index, message.data.clone()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::charts::IWorkChartArchive;
    use crate::wire::append_length_delimited_field;

    fn encoded_chart(frame: Option<tsch::RectArchive>) -> Vec<u8> {
        IWorkChartArchive::new(
            tsch::ChartDrawableArchive::default(),
            tsch::ChartArchive {
                legend_frame: frame,
                ..Default::default()
            },
        )
        .encode()
        .unwrap()
    }

    #[test]
    fn legend_frame_is_strict_and_wire_exact() {
        let baseline = encoded_chart(None);
        assert_eq!(
            read_chart_legend_frame(&baseline).unwrap(),
            ChartLegendFrame::Automatic
        );
        let rect = ChartLegendRect::from_points(-4.5, 12.25, 0.0, 38.0).unwrap();
        let chart = chart_payload(&baseline).unwrap();
        let mut chart_with_unknown = chart.to_vec();
        append_length_delimited_field(&mut chart_with_unknown, 9_999, b"future").unwrap();
        let with_unknown = patch_length_delimited_field(
            &baseline,
            CHART_ARCHIVE_EXTENSION_FIELD,
            true,
            Some(&chart_with_unknown),
        )
        .unwrap();
        let frame = rect_to_native(rect).encode_to_vec();
        let chart = patch_length_delimited_field(
            &chart_with_unknown,
            CHART_LEGEND_FRAME_FIELD,
            false,
            Some(&frame),
        )
        .unwrap();
        let framed = patch_length_delimited_field(
            &with_unknown,
            CHART_ARCHIVE_EXTENSION_FIELD,
            true,
            Some(&chart),
        )
        .unwrap();
        assert_eq!(
            read_chart_legend_frame(&framed).unwrap(),
            ChartLegendFrame::Frame(rect)
        );
        let reset_chart = patch_length_delimited_field(
            chart_payload(&framed).unwrap(),
            CHART_LEGEND_FRAME_FIELD,
            true,
            None,
        )
        .unwrap();
        let reset = patch_length_delimited_field(
            &framed,
            CHART_ARCHIVE_EXTENSION_FIELD,
            true,
            Some(&reset_chart),
        )
        .unwrap();
        assert_eq!(reset, with_unknown);
    }

    #[test]
    fn invalid_geometry_and_malformed_wire_are_rejected() {
        assert!(ChartLegendCoordinate::new(f32::NAN).is_err());
        assert!(ChartLegendExtent::new(f32::INFINITY).is_err());
        assert!(ChartLegendExtent::new(-1.0).is_err());

        let rect = rect_to_native(ChartLegendRect::from_points(1.0, 2.0, 3.0, 4.0).unwrap())
            .encode_to_vec();
        let baseline = encoded_chart(None);
        let chart = chart_payload(&baseline).unwrap();
        let once =
            patch_length_delimited_field(chart, CHART_LEGEND_FRAME_FIELD, false, Some(&rect))
                .unwrap();
        let mut duplicate = once;
        append_length_delimited_field(&mut duplicate, CHART_LEGEND_FRAME_FIELD, &rect).unwrap();
        let malformed = patch_length_delimited_field(
            &baseline,
            CHART_ARCHIVE_EXTENSION_FIELD,
            true,
            Some(&duplicate),
        )
        .unwrap();
        assert!(read_chart_legend_frame(&malformed).is_err());
    }
}
