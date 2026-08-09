//! BIFF wire identifiers and bounded primitive codecs for the chart owner.

use litchi_biff::{Encoder as GraphEncoder, Kind as RecordKind, Limits as BiffLimits};
use litchi_ograph::Limits as GraphLimits;
use litchi_ograph::chart::format;
use litchi_ograph::record::{chart3d, frame, line, marker, pie, series};

use super::model::{AreaFormat, Limits, LineFormat};
use crate::{Error, Result};

pub(crate) const BOF: u16 = 0x0809;
pub(crate) const EOF: u16 = 0x000a;
pub(crate) const CONTINUE: u16 = 0x003c;
pub(crate) const OBJ: u16 = 0x005d;
pub(crate) const BOUNDSHEET: u16 = 0x0085;
pub(crate) const WINDOW1: u16 = 0x003d;
pub(crate) const RR_TAB_ID: u16 = 0x013d;
pub(crate) const SUP_BOOK: u16 = 0x01ae;
pub(crate) const EXTERN_SHEET: u16 = 0x0017;
pub(crate) const LBL: u16 = 0x0018;
pub(crate) const BLANK: u16 = 0x0201;
pub(crate) const NUMBER: u16 = 0x0203;
pub(crate) const LABEL: u16 = 0x0204;
pub(crate) const CHART: u16 = 0x1002;
pub(crate) const SERIES: u16 = 0x1003;
pub(crate) const BRAI: u16 = 0x1051;
pub(crate) const DATA_FORMAT: u16 = 0x1006;
pub(crate) const LINE_FORMAT: u16 = 0x1007;
pub(crate) const MARKER_FORMAT: u16 = 0x1009;
pub(crate) const AREA_FORMAT: u16 = 0x100a;
pub(crate) const PIE_FORMAT: u16 = pie::Format::KIND.get();
pub(crate) const SERIES_TEXT: u16 = 0x100d;
pub(crate) const CHART_FORMAT: u16 = 0x1014;
pub(crate) const LEGEND: u16 = 0x1015;
pub(crate) const SERIES_LIST: u16 = 0x1016;
pub(crate) const BAR: u16 = 0x1017;
pub(crate) const LINE: u16 = 0x1018;
pub(crate) const PIE: u16 = 0x1019;
pub(crate) const AREA: u16 = 0x101a;
pub(crate) const SCATTER: u16 = 0x101b;
pub(crate) const CRT_LINE: u16 = line::Line::KIND.get();
pub(crate) const CRT_LINK: u16 = line::Link::KIND.get();
pub(crate) const AXIS: u16 = 0x101d;
pub(crate) const TICK: u16 = 0x101e;
pub(crate) const VALUE_RANGE: u16 = 0x101f;
pub(crate) const CAT_SER_RANGE: u16 = 0x1020;
pub(crate) const AXIS_LINE: u16 = 0x1021;
pub(crate) const DEFAULT_TEXT: u16 = 0x1024;
pub(crate) const TEXT: u16 = 0x1025;
pub(crate) const FONT_X: u16 = 0x1026;
pub(crate) const OBJECT_LINK: u16 = 0x1027;
pub(crate) const FRAME: u16 = frame::Frame::KIND.get();
pub(crate) const BEGIN: u16 = marker::Begin::KIND.get();
pub(crate) const END: u16 = marker::End::KIND.get();
pub(crate) const PLOT_AREA: u16 = marker::PlotArea::KIND.get();
pub(crate) const DROP_BAR: u16 = 0x103d;
pub(crate) const RADAR: u16 = 0x103e;
pub(crate) const SURFACE: u16 = 0x103f;
pub(crate) const RADAR_AREA: u16 = 0x1040;
pub(crate) const AXIS_PARENT: u16 = 0x1041;
pub(crate) const SHT_PROPS: u16 = 0x1044;
pub(crate) const SER_TO_CRT: u16 = 0x1045;
pub(crate) const AXES_USED: u16 = 0x1046;
pub(crate) const SERIES_PARENT: u16 = series::Parent::KIND.get();
pub(crate) const SERIES_FORMAT: u16 = series::Format::KIND.get();
pub(crate) const BAR_SHAPE: u16 = chart3d::BarShape::KIND.get();
pub(crate) const DATA_LAB_EXT: u16 = 0x086a;
pub(crate) const DATA_LAB_EXT_CONTENTS: u16 = 0x086b;
pub(crate) const PLOT_GROWTH: u16 = 0x1064;
pub(crate) const SI_INDEX: u16 = 0x1065;
pub(crate) fn parse_line_format(data: &[u8]) -> Result<LineFormat> {
    exact(data, 12, LINE_FORMAT)?;
    Ok(LineFormat {
        color: array_at(data, 0)?,
        pattern: u16_at(data, 4)?,
        weight: i16_at(data, 6)?,
        flags: u16_at(data, 8)?,
        color_index: u16_at(data, 10)?,
    })
}
#[cfg(test)]
pub(crate) fn write_line(v: &LineFormat) -> Vec<u8> {
    let mut d = v.color.to_vec();
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.weight.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.color_index.to_le_bytes());
    d
}
pub(crate) fn shared_line(value: &LineFormat) -> format::Line {
    format::Line {
        color: value.color,
        pattern: value.pattern,
        weight: value.weight,
        flags: value.flags,
        color_index: value.color_index,
    }
}
#[cfg(test)]
pub(crate) fn shared_line_bytes(value: &format::Line) -> Vec<u8> {
    write_line(&LineFormat {
        color: value.color,
        pattern: value.pattern,
        weight: value.weight,
        flags: value.flags,
        color_index: value.color_index,
    })
}
pub(crate) fn parse_area_format(data: &[u8]) -> Result<AreaFormat> {
    exact(data, 16, AREA_FORMAT)?;
    Ok(AreaFormat {
        foreground: array_at(data, 0)?,
        background: array_at(data, 4)?,
        pattern: u16_at(data, 8)?,
        flags: u16_at(data, 10)?,
        foreground_index: u16_at(data, 12)?,
        background_index: u16_at(data, 14)?,
    })
}
#[cfg(test)]
pub(crate) fn write_area(v: &AreaFormat) -> Vec<u8> {
    let mut d = v.foreground.to_vec();
    d.extend(v.background);
    d.extend(v.pattern.to_le_bytes());
    d.extend(v.flags.to_le_bytes());
    d.extend(v.foreground_index.to_le_bytes());
    d.extend(v.background_index.to_le_bytes());
    d
}
pub(crate) fn shared_area(value: &AreaFormat) -> format::Area {
    format::Area {
        foreground: value.foreground,
        background: value.background,
        pattern: value.pattern,
        flags: value.flags,
        foreground_index: value.foreground_index,
        background_index: value.background_index,
    }
}
#[cfg(test)]
pub(crate) fn shared_area_bytes(value: &format::Area) -> Vec<u8> {
    write_area(&AreaFormat {
        foreground: value.foreground,
        background: value.background,
        pattern: value.pattern,
        flags: value.flags,
        foreground_index: value.foreground_index,
        background_index: value.background_index,
    })
}
pub(crate) fn parse_short_text(data: &[u8]) -> Result<String> {
    if data.len() < 4 || u16_at(data, 0)? != 0 {
        return invalid(
            SERIES_TEXT,
            "SeriesText is truncated or reserved field is nonzero",
        );
    }
    parse_biff8_string(&data[2..])
}
#[cfg(test)]
pub(crate) fn short_text(value: &str) -> Result<Vec<u8>> {
    let mut d = 0u16.to_le_bytes().to_vec();
    d.extend(biff8_string(value)?);
    Ok(d)
}
pub(crate) fn parse_biff8_string(data: &[u8]) -> Result<String> {
    if data.len() < 2 {
        return invalid(SERIES_TEXT, "chart string is truncated");
    }
    let count = usize::from(data[0]);
    let wide = data[1] & 1 != 0;
    if data[1] & !1 != 0 {
        return invalid(SERIES_TEXT, "chart string uses unsupported option flags");
    }
    let need = 2 + count * (if wide { 2 } else { 1 });
    if data.len() != need {
        return invalid(SERIES_TEXT, "chart string length mismatch");
    }
    if wide {
        let units = data[2..]
            .chunks_exact(2)
            .map(|v| u16::from_le_bytes([v[0], v[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_error| invalid_error(SERIES_TEXT, "invalid UTF-16 chart string"))
    } else {
        Ok(data[2..].iter().map(|v| char::from(*v)).collect())
    }
}

/// Parses the `XLUnicodeString` used by an XLS `Label` cache record.
pub(crate) fn parse_xl_unicode_string(data: &[u8]) -> Result<String> {
    if data.len() < 3 {
        return invalid(LABEL, "cached Label string is truncated");
    }
    let count = usize::from(u16_at(data, 0)?);
    let flags = data[2];
    if flags & !1 != 0 {
        return invalid(LABEL, "cached Label string uses reserved flags");
    }
    let wide = flags & 1 != 0;
    let width = if wide { 2 } else { 1 };
    let bytes = count
        .checked_mul(width)
        .ok_or_else(|| Error::InvalidData("cached Label string length overflow".into()))?;
    let end = 3usize
        .checked_add(bytes)
        .ok_or_else(|| Error::InvalidData("cached Label string length overflow".into()))?;
    if data.len() != end {
        return invalid(LABEL, "cached Label string length mismatch");
    }
    if wide {
        let units = data[3..]
            .chunks_exact(2)
            .map(|value| u16::from_le_bytes([value[0], value[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_error| invalid_error(LABEL, "cached Label string is invalid UTF-16"))
    } else {
        Ok(data[3..].iter().map(|value| char::from(*value)).collect())
    }
}

/// Encodes one bounded `XLUnicodeString` for an XLS `Label` cache record.
pub(crate) fn xl_unicode_string(value: &str) -> Result<Vec<u8>> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > 255 {
        return invalid(LABEL, "cached Label string exceeds 255 UTF-16 code units");
    }
    let wide = units.iter().any(|value| *value > 255);
    let mut data = Vec::with_capacity(3 + units.len() * if wide { 2 } else { 1 });
    data.extend(crate::utils::truncate_usize_to_u16(units.len()).to_le_bytes());
    data.push(u8::from(wide));
    if wide {
        for value in units {
            data.extend(value.to_le_bytes());
        }
    } else {
        data.extend(units.into_iter().map(crate::utils::truncate_u16_to_u8));
    }
    Ok(data)
}

#[cfg(test)]
pub(crate) fn biff8_string(value: &str) -> Result<Vec<u8>> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.len() > 255 {
        return invalid(SERIES_TEXT, "chart string exceeds 255 UTF-16 code units");
    }
    let wide = units.iter().any(|v| *v > 255);
    let mut d = vec![units.len() as u8, u8::from(wide)];
    if wide {
        for v in units {
            d.extend(v.to_le_bytes());
        }
    } else {
        d.extend(units.into_iter().map(|v| v as u8));
    }
    Ok(d)
}
pub(crate) fn chart_scan_limits(limits: Limits) -> GraphLimits {
    GraphLimits {
        max_workbook_bytes: limits.max_workbook_bytes,
        max_charts: limits.max_charts,
        max_chart_records: limits.max_records_per_chart,
        max_series: limits.max_series,
        max_groups: limits.max_groups,
        max_axes: limits.max_axes,
        max_formula_bytes: limits.max_formula_bytes,
        max_cached_values: limits.max_cached_values,
        max_unknown_bytes: limits.max_unknown_bytes,
        biff: BiffLimits {
            max_records: limits.max_records_per_chart,
            max_input_bytes: limits.max_workbook_bytes,
            max_output_bytes: limits.max_workbook_bytes,
            ..BiffLimits::default()
        },
        ..GraphLimits::default()
    }
}
#[cfg(test)]
pub(crate) fn chart_encoder(limits: Limits) -> Result<GraphEncoder> {
    GraphEncoder::with_limits(BiffLimits {
        max_records: limits.max_records_per_chart,
        max_output_bytes: limits.max_workbook_bytes,
        ..BiffLimits::default()
    })
    .map_err(|error| frame_error(CHART, error))
}
pub(crate) fn push_record(out: &mut GraphEncoder, kind: u16, data: &[u8]) -> Result<()> {
    out.push(RecordKind::from_wire(kind), data)
        .map_err(|error| frame_error(kind, error))
}
pub(crate) fn record(kind: u16, data: &[u8]) -> Result<Vec<u8>> {
    let output_bytes = data
        .len()
        .checked_add(4)
        .ok_or_else(|| Error::InvalidData("BIFF record size overflow".into()))?;
    let limits = BiffLimits {
        max_records: 1,
        max_output_bytes: output_bytes.max(1),
        ..BiffLimits::default()
    };
    let mut out = GraphEncoder::with_limits(limits).map_err(|error| frame_error(kind, error))?;
    push_record(&mut out, kind, data)?;
    Ok(out.finish())
}
#[cfg(test)]
pub(crate) fn known_record(kind: u16) -> bool {
    matches!(
        kind,
        BOF | EOF
            | CHART
            | SERIES
            | BRAI
            | SER_TO_CRT
            | SERIES_TEXT
            | CHART_FORMAT
            | BAR
            | LINE
            | PIE
            | AREA
            | SCATTER
            | RADAR
            | RADAR_AREA
            | SURFACE
            | CRT_LINE
            | DROP_BAR
            | AXIS
            | VALUE_RANGE
            | TICK
            | AXIS_LINE
            | LINE_FORMAT
            | AREA_FORMAT
            | MARKER_FORMAT
            | DATA_FORMAT
            | PIE_FORMAT
            | LEGEND
            | PLOT_AREA
            | DATA_LAB_EXT
            | DATA_LAB_EXT_CONTENTS
            | TEXT
            | SI_INDEX
            | BLANK
            | NUMBER
            | LABEL
            | BEGIN
            | END
            | SHT_PROPS
            | AXES_USED
            | AXIS_PARENT
    )
}
pub(crate) fn validate_limits(v: Limits) -> Result<()> {
    if v.max_workbook_bytes == 0
        || v.max_charts == 0
        || v.max_records_per_chart == 0
        || v.max_series == 0
        || v.max_groups == 0
        || v.max_axes == 0
        || v.max_formula_bytes == 0
        || v.max_cached_values == 0
        || v.max_unknown_bytes == 0
    {
        return Err(Error::InvalidData(
            "all chart limits must be nonzero".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_sheet_properties(flags: u32) -> Result<()> {
    let blank = (flags >> 16) & 0xff;
    let always_auto = flags & (1 << 4) != 0;
    let manual_plot = flags & (1 << 3) != 0;
    if flags & 0xff00_ffe0 != 0 || blank > 2 || (always_auto && !manual_plot) {
        return invalid(
            SHT_PROPS,
            "ShtProps reserved bits, blank mode, or plot-area flags are invalid",
        );
    }
    Ok(())
}
pub(crate) fn bounded_count(data: &[u8], offset: usize) -> Result<u16> {
    let v = u16_at(data, offset)?;
    if v > 32767 {
        return invalid(SERIES, "series value count exceeds 32767");
    }
    Ok(v)
}
pub(crate) fn exact(data: &[u8], len: usize, kind: u16) -> Result<()> {
    if data.len() != len {
        return invalid(kind, format!("record must contain {len} bytes"));
    }
    Ok(())
}
pub(crate) fn u16_at(data: &[u8], o: usize) -> Result<u16> {
    Ok(u16::from_le_bytes(array_at(data, o)?))
}
pub(crate) fn i16_at(data: &[u8], o: usize) -> Result<i16> {
    Ok(crate::utils::wrap_u16_to_i16(u16_at(data, o)?))
}
pub(crate) fn u32_at(data: &[u8], o: usize) -> Result<u32> {
    Ok(u32::from_le_bytes(array_at(data, o)?))
}
pub(crate) fn i32_at(data: &[u8], o: usize) -> Result<i32> {
    Ok(crate::utils::wrap_u32_to_i32(u32_at(data, o)?))
}
pub(crate) fn f64_at(data: &[u8], o: usize) -> Result<f64> {
    Ok(f64::from_le_bytes(array_at(data, o)?))
}
pub(crate) fn array_at<const N: usize>(data: &[u8], offset: usize) -> Result<[u8; N]> {
    let expected = offset
        .checked_add(N)
        .ok_or_else(|| Error::InvalidData("record field offset overflow".into()))?;
    data.get(offset..expected)
        .and_then(|bytes| bytes.try_into().ok())
        .ok_or(Error::InvalidLength {
            expected,
            found: data.len(),
        })
}
pub(crate) fn invalid_error(kind: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: kind,
        message: message.into(),
    }
}
pub(crate) fn graph_error(kind: u16, error: litchi_ograph::Error) -> Error {
    invalid_error(kind, error.to_string())
}
pub(crate) fn frame_error(kind: u16, error: litchi_biff::Error) -> Error {
    invalid_error(kind, error.to_string())
}
pub(crate) fn invalid<T>(kind: u16, message: impl Into<String>) -> Result<T> {
    Err(invalid_error(kind, message))
}
