//! MS-OGRAPH chart-substream emission.

use super::super::super::axis::{self, Axis};
use super::super::super::cache as chart_cache;
use super::super::super::format::{self, Format};
use super::super::super::layout;
use super::super::super::model::{Cache, Chart, DataKind, Family, Group, Owner};
use super::super::super::{
    BOF, BOF_BYTES, EOF, EXCEL_DOC_TYPE, EXCEL_VERSION, GRAPH_DOC_TYPE, GRAPH_VERSION,
    Kind as ChartKind,
};
use super::cache;
use super::links;
use super::text::short_text;
use super::validate;
use super::wire::{
    AREA, AREA_FORMAT, AXES_USED, AXIS, AXIS_LINE, AXIS_PARENT, BAR, BEGIN, BRAI, CHART_FORMAT,
    CHART_REC, CRT_LINE, CRT_LINK, DATA_FORMAT, DIMENSIONS, DROP_BAR, END, LEGEND, LINE,
    LINE_FORMAT, MARKER_FORMAT, PIE, PIE_FORMAT, PLOT_AREA, PLOT_GROWTH, POS, RADAR, RADAR_AREA,
    SCATTER, SCL, SER_AUX_ERR_BAR, SER_AUX_TREND, SER_TO_CRT, SERIES, SERIES_TEXT, SHT_PROPS,
    SI_INDEX, SURFACE, TICK, VALUE_RANGE, push_record, put_byte, put_f64, put_i16, put_i32,
    put_slice, put_u16, put_u32,
};
use crate::{Error, Limits, Result};
use litchi_biff::Encoder;

pub(crate) fn encode(chart: &Chart, limits: Limits) -> Result<Vec<u8>> {
    if !chart.authoring_proven {
        return Err(Error::UnsupportedAuthoring {
            reason: "the host prelude, attached-label/frame grammar, and complete axis ownership are not yet proven",
        });
    }
    if chart.parents.len() != 1 || !(1..=4).contains(&chart.groups.len()) {
        return Err(Error::UnsupportedAuthoring {
            reason: "secondary-axis chart-group ownership is not yet modeled",
        });
    }
    validate::validate(chart, limits, true)?;
    let mut out = Encoder::with_limits(limits.biff)?;
    push_record(&mut out, BOF, &bof(chart.context.kind()))?;

    let mut rect = [0u8; 16];
    rect.get_mut(0..4)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.x.to_le_bytes());
    rect.get_mut(4..8)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.y.to_le_bytes());
    rect.get_mut(8..12)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.width.to_le_bytes());
    rect.get_mut(12..16)
        .ok_or(Error::SizeOverflow {
            resource: "Chart rectangle",
        })?
        .copy_from_slice(&chart.rect.height.to_le_bytes());
    push_record(&mut out, CHART_REC, &rect)?;
    push_record(&mut out, BEGIN, &[])?;
    let mut zoom = [0u8; 4];
    put_u16(&mut zoom, 0, chart.zoom.numerator())?;
    put_u16(&mut zoom, 2, chart.zoom.denominator())?;
    push_record(&mut out, SCL, &zoom)?;
    let mut growth = [0u8; 8];
    put_i32(&mut growth, 0, chart.growth.x.raw())?;
    put_i32(&mut growth, 4, chart.growth.y.raw())?;
    push_record(&mut out, PLOT_GROWTH, &growth)?;

    for series in &chart.series {
        let mut data = [0u8; 12];
        put_u16(
            &mut data,
            0,
            match series.category_kind {
                DataKind::Numeric => 1,
                DataKind::Text => 3,
            },
        )?;
        put_u16(&mut data, 2, 1)?;
        put_u16(&mut data, 4, series.category_count.get())?;
        put_u16(&mut data, 6, series.value_count.get())?;
        put_u16(&mut data, 8, 1)?;
        put_u16(&mut data, 10, series.bubble_count.get())?;
        push_record(&mut out, SERIES, &data)?;
        push_record(&mut out, BEGIN, &[])?;
        for binding in series.ai.ordered() {
            let data = links::encode_link(binding.link(), chart.context, limits)?;
            push_record(&mut out, BRAI, &data)?;
            if let Some(text) = binding.text() {
                push_record(&mut out, SERIES_TEXT, &short_text(text)?)?;
            }
        }
        match &series.owner {
            Owner::Group(group) => {
                push_record(&mut out, SER_TO_CRT, &u16::from(group.get()).to_le_bytes())?;
            },
            Owner::Trend { parent, data } => {
                parent.write(&mut out)?;
                push_record(&mut out, SER_AUX_TREND, data)?;
            },
            Owner::ErrorBar { parent, data } => {
                parent.write(&mut out)?;
                push_record(&mut out, SER_AUX_ERR_BAR, data)?;
            },
        }
        push_record(&mut out, END, &[])?;
    }

    push_record(&mut out, SHT_PROPS, &chart.props.flags.to_le_bytes())?;
    push_record(
        &mut out,
        AXES_USED,
        &u16::try_from(chart.parents.len())
            .ok()
            .ok_or(Error::SizeOverflow {
                resource: "axis-parent count",
            })?
            .to_le_bytes(),
    )?;
    let parent = chart.parents.first().copied().ok_or(Error::InvalidModel {
        field: "axis parents",
        reason: "a chart requires a primary AxisParent",
    })?;
    let mut parent_body = [0u8; 18];
    put_u16(&mut parent_body, 0, u16::from(parent.is_secondary()))?;
    push_record(&mut out, AXIS_PARENT, &parent_body)?;
    push_record(&mut out, BEGIN, &[])?;
    encode_pos(&mut out, parent.pos())?;
    for axis in &chart.axes {
        encode_axis(&mut out, axis)?;
    }
    for group in &chart.groups {
        encode_group(&mut out, group)?;
    }
    if chart.props.plot_area {
        push_record(&mut out, PLOT_AREA, &[])?;
    }
    if let Some(legend) = chart.legend {
        let mut data = [0u8; 20];
        put_i32(&mut data, 0, legend.x)?;
        put_i32(&mut data, 4, legend.y)?;
        put_i32(&mut data, 8, legend.width)?;
        put_i32(&mut data, 12, legend.height)?;
        put_byte(&mut data, 16, legend.position)?;
        put_byte(&mut data, 17, legend.spacing)?;
        put_u16(&mut data, 18, legend.flags)?;
        push_record(&mut out, LEGEND, &data)?;
    }
    if let Some(title) = &chart.title {
        push_record(&mut out, SERIES_TEXT, &short_text(title)?)?;
    }
    for value in &chart.formats {
        encode_format(&mut out, value)?;
    }
    for label in &chart.labels {
        push_record(&mut out, label.kind, &label.data)?;
    }

    push_record(&mut out, END, &[])?;
    push_record(&mut out, END, &[])?;
    encode_dimensions(&mut out, chart.dimensions)?;
    match chart.context.kind() {
        ChartKind::Excel => {
            for index in chart_cache::Index::ALL {
                push_record(&mut out, SI_INDEX, &(index as u16).to_le_bytes())?;
                for value in chart.caches.iter().filter(
                    |value| matches!(value, Cache::Excel { section, .. } if *section == index),
                ) {
                    cache::encode_cache(&mut out, value)?;
                }
            }
        },
        ChartKind::Graph => {
            for value in &chart.caches {
                cache::encode_cache(&mut out, value)?;
            }
        },
    }
    push_record(&mut out, EOF, &[])?;
    Ok(out.finish())
}

fn encode_pos(out: &mut Encoder, pos: layout::Pos) -> Result<()> {
    let mut data = [0u8; 20];
    put_u16(&mut data, 0, pos.top_left() as u16)?;
    put_u16(&mut data, 2, pos.bottom_right() as u16)?;
    put_i16(&mut data, 4, pos.x())?;
    put_i16(&mut data, 8, pos.y())?;
    put_i16(&mut data, 12, pos.width())?;
    put_i16(&mut data, 16, pos.height())?;
    push_record(out, POS, &data)
}

fn encode_dimensions(out: &mut Encoder, dimensions: chart_cache::Dims) -> Result<()> {
    let mut data = [0u8; 14];
    match dimensions {
        chart_cache::Dims::Excel(value) => {
            put_u32(&mut data, 0, value.first_row())?;
            put_u32(&mut data, 4, value.row_after())?;
            put_u16(&mut data, 8, value.first_col())?;
            put_u16(&mut data, 10, value.col_after())?;
        },
        chart_cache::Dims::Graph(value) => {
            put_u32(&mut data, 4, u32::from(value.longest_row().get()))?;
            put_u16(&mut data, 10, u16::from(value.rows()))?;
        },
    }
    push_record(out, DIMENSIONS, &data)
}

fn bof(kind: ChartKind) -> [u8; BOF_BYTES] {
    let mut data = [0u8; BOF_BYTES];
    match kind {
        ChartKind::Excel => {
            data[0..2].copy_from_slice(&EXCEL_VERSION.to_le_bytes());
            data[2..4].copy_from_slice(&EXCEL_DOC_TYPE.to_le_bytes());
            data[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
            data[6..8].copy_from_slice(&0x07CC_u16.to_le_bytes());
            data[8..12].copy_from_slice(&0x0000_0009_u32.to_le_bytes());
            data[12..16].copy_from_slice(&6u32.to_le_bytes());
        },
        ChartKind::Graph => {
            data[0..2].copy_from_slice(&GRAPH_VERSION.to_le_bytes());
            data[2..4].copy_from_slice(&GRAPH_DOC_TYPE.to_le_bytes());
            data[4..6].copy_from_slice(&0x0DBB_u16.to_le_bytes());
            data[6..8].copy_from_slice(&0x07CD_u16.to_le_bytes());
            data[8..12].copy_from_slice(&(0x0000_0009_u32 | (6 << 14)).to_le_bytes());
            data[12..16].copy_from_slice(&(0x06_u32 | (6 << 8)).to_le_bytes());
        },
    }
    data
}

fn encode_axis(out: &mut Encoder, axis: &Axis) -> Result<()> {
    let mut body = [0u8; 18];
    put_u16(
        &mut body,
        0,
        match axis.kind {
            axis::Kind::Category => 0,
            axis::Kind::Value => 1,
            axis::Kind::Series => 2,
        },
    )?;
    push_record(out, AXIS, &body)?;
    push_record(out, BEGIN, &[])?;
    if let Some(scale) = axis.scale {
        let mut data = [0u8; 42];
        put_f64(&mut data, 0, scale.min)?;
        put_f64(&mut data, 8, scale.max)?;
        put_f64(&mut data, 16, scale.major)?;
        put_f64(&mut data, 24, scale.minor)?;
        put_f64(&mut data, 32, scale.crossing)?;
        put_u16(&mut data, 40, scale.flags)?;
        push_record(out, VALUE_RANGE, &data)?;
    }
    if let Some(tick) = axis.tick {
        let mut data = [0u8; 26];
        put_byte(&mut data, 0, tick.major)?;
        put_byte(&mut data, 1, tick.minor)?;
        put_byte(&mut data, 2, tick.label)?;
        put_byte(&mut data, 3, tick.background)?;
        put_slice(&mut data, 4, &tick.color, "Tick color")?;
        put_u16(&mut data, 24, tick.flags)?;
        push_record(out, TICK, &data)?;
    }
    for line in &axis.lines {
        push_record(
            out,
            AXIS_LINE,
            &u16::from(validate::line_kind(line.kind)).to_le_bytes(),
        )?;
        push_record(out, LINE_FORMAT, &line_bytes(line.format))?;
    }
    push_record(out, END, &[])
}

fn encode_group(out: &mut Encoder, group: &Group) -> Result<()> {
    let mut header = [0u8; 20];
    put_u16(&mut header, 16, u16::from(group.vary_colors))?;
    put_u16(&mut header, 18, u16::from(group.order.get()))?;
    push_record(out, CHART_FORMAT, &header)?;
    push_record(out, BEGIN, &[])?;
    match group.family {
        Family::Line { flags } => push_record(out, LINE, &flags.to_le_bytes())?,
        Family::Area { flags } => push_record(out, AREA, &flags.to_le_bytes())?,
        Family::Bar {
            overlap,
            gap,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_i16(&mut data, 0, overlap.get())?;
            put_u16(&mut data, 2, gap.get())?;
            put_u16(&mut data, 4, flags)?;
            push_record(out, BAR, &data)?;
        },
        Family::Pie {
            rotation,
            hole,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, rotation)?;
            put_u16(&mut data, 2, hole)?;
            put_u16(&mut data, 4, flags)?;
            push_record(out, PIE, &data)?;
        },
        Family::Scatter {
            bubble_percent,
            bubble_kind,
            flags,
        } => {
            let mut data = [0u8; 6];
            put_u16(&mut data, 0, bubble_percent.get())?;
            put_u16(&mut data, 2, bubble_kind as u16)?;
            put_u16(&mut data, 4, flags)?;
            push_record(out, SCATTER, &data)?;
        },
        Family::Radar { filled, flags } => {
            push_record(
                out,
                if filled { RADAR_AREA } else { RADAR },
                &flags.to_le_bytes(),
            )?;
        },
        Family::Surface { flags } => push_record(out, SURFACE, &flags.to_le_bytes())?,
    }
    push_record(out, CRT_LINK, &group.link.payload())?;
    for line in &group.lines {
        let marker = crate::record::line::Line::new(line.kind);
        push_record(out, CRT_LINE, &marker.payload())?;
        push_record(out, LINE_FORMAT, &line_bytes(line.format))?;
    }
    for drop in &group.drop_bars {
        push_record(out, DROP_BAR, &drop.gap.get().to_le_bytes())?;
        push_record(out, BEGIN, &[])?;
        push_record(out, LINE_FORMAT, &line_bytes(drop.line))?;
        push_record(out, AREA_FORMAT, &area_bytes(drop.area))?;
        push_record(out, END, &[])?;
    }
    push_record(out, END, &[])
}

fn encode_format(out: &mut Encoder, value: &Format) -> Result<()> {
    match value {
        Format::Line(value) => push_record(out, LINE_FORMAT, &line_bytes(*value)),
        Format::Area(value) => push_record(out, AREA_FORMAT, &area_bytes(*value)),
        Format::Marker { data } => push_record(out, MARKER_FORMAT, data),
        Format::Data {
            point,
            series,
            flags,
        } => {
            let mut data = [0u8; 8];
            put_u16(&mut data, 0, *point)?;
            put_u16(&mut data, 2, *series)?;
            put_u16(&mut data, 4, 0)?;
            put_u16(&mut data, 6, *flags)?;
            push_record(out, DATA_FORMAT, &data)
        },
        Format::Pie { explosion } => push_record(out, PIE_FORMAT, &explosion.to_le_bytes()),
    }
}

fn line_bytes(value: format::Line) -> [u8; 12] {
    let mut data = [0u8; 12];
    data[0..4].copy_from_slice(&value.color);
    data[4..6].copy_from_slice(&value.pattern.to_le_bytes());
    data[6..8].copy_from_slice(&value.weight.to_le_bytes());
    data[8..10].copy_from_slice(&value.flags.to_le_bytes());
    data[10..12].copy_from_slice(&value.color_index.to_le_bytes());
    data
}

fn area_bytes(value: format::Area) -> [u8; 16] {
    let mut data = [0u8; 16];
    data[0..4].copy_from_slice(&value.foreground);
    data[4..8].copy_from_slice(&value.background);
    data[8..10].copy_from_slice(&value.pattern.to_le_bytes());
    data[10..12].copy_from_slice(&value.flags.to_le_bytes());
    data[12..14].copy_from_slice(&value.foreground_index.to_le_bytes());
    data[14..16].copy_from_slice(&value.background_index.to_le_bytes());
    data
}
